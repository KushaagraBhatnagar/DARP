import { withRoute, fail, ok, file } from '@/server/http/handler';
import { computeNbaReport } from '@/server/reports/nba';
import { generateNbaExcelWorkbook } from '@/server/export/nba-workbook';
import { activeCycle, ServiceError } from '@/server/records/service';
import { db } from '@/server/db';
import { departments } from '@/server/db/schema';

export const GET = withRoute(
  { roles: ['admin', 'hod', 'dofa', 'drie'], rate: { limit: 30, windowSeconds: 60 } },
  async ({ actor, query: sp }) => {
    let cycleId = sp.get('cycleId');
    if (!cycleId) {
      try {
        cycleId = (await activeCycle()).id;
      } catch (err) {
        if (err instanceof ServiceError) return fail(err.code, { message: err.message });
        throw err;
      }
    }

    // Determine department
    let departmentId = sp.get('departmentId');
    if (actor.role === 'hod') {
      departmentId = actor.departmentId ?? null;
    }
    if (!departmentId) {
      // Pick first active department if not specified
      const [firstDept] = await db.select().from(departments).limit(1);
      if (!firstDept) return fail('NOT_FOUND', { message: 'No department found' });
      departmentId = firstDept.id;
    }

    const report = await computeNbaReport(departmentId, cycleId);

    // If download requested
    if (sp.get('format') === 'xlsx' || sp.get('download') === '1') {
      const wb = await generateNbaExcelWorkbook(report);
      const buffer = await wb.xlsx.writeBuffer();
      const safeCycle = report.cycle.name.replace(/[^a-zA-Z0-9_-]/g, '_');
      const filename = `NBA_Tier1_SAR_Criterion5_${report.department.code}_${safeCycle}.xlsx`;

      return file(new Uint8Array(buffer), {
        'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'content-disposition': `attachment; filename="${filename}"`,
      });
    }

    // Otherwise return JSON
    return ok(report);
  }
);
