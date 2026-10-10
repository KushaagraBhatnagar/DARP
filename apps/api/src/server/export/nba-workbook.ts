import ExcelJS from 'exceljs';
import type { NbaCriterion5Report } from '@/server/reports/nba';

/**
 * Builds an official NBA Tier-1 SAR Criterion 5 Workbook with formatting and summary sheets.
 */
export async function generateNbaExcelWorkbook(
  report: NbaCriterion5Report
): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'DARP · Birla Institute of Technology, Mesra';
  wb.lastModifiedBy = 'IQAC Administrator';
  wb.created = new Date();
  wb.modified = new Date();

  const brandColor = '1E3A8A'; // Deep Blue
  const goldColor = 'D97706'; // Gold Accent
  const headerFill: ExcelJS.FillPattern = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: brandColor },
  };
  const subHeaderFill: ExcelJS.FillPattern = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'F1F5F9' },
  };

  // ─────────────────────────────────────────────────────────────
  // SHEET 1: EXECUTIVE SUMMARY & SCORECARD
  // ─────────────────────────────────────────────────────────────
  const summarySheet = wb.addWorksheet('Criterion 5 Summary', {
    views: [{ showGridLines: true }],
  });

  summarySheet.columns = [
    { width: 12 },
    { width: 38 },
    { width: 14 },
    { width: 16 },
    { width: 50 },
  ];

  summarySheet.addRow(['NBA TIER-1 SELF ASSESSMENT REPORT (SAR)']);
  summarySheet.addRow(['CRITERION 5: FACULTY INFORMATION AND CONTRIBUTIONS (200 MARKS)']);
  summarySheet.addRow([`Department: ${report.department.name} (${report.department.code})`]);
  summarySheet.addRow([`Accreditation Cycle: ${report.cycle.name} (${report.cycle.startYear} - ${report.cycle.endYear})`]);
  summarySheet.addRow([]);

  // Title formatting
  summarySheet.mergeCells('A1:E1');
  summarySheet.mergeCells('A2:E2');
  summarySheet.mergeCells('A3:E3');
  summarySheet.mergeCells('A4:E4');

  summarySheet.getCell('A1').font = { bold: true, size: 14, color: { argb: 'FFFFFF' } };
  summarySheet.getCell('A1').fill = headerFill;
  summarySheet.getCell('A1').alignment = { horizontal: 'center' };

  summarySheet.getCell('A2').font = { bold: true, size: 11, color: { argb: 'FFFFFF' } };
  summarySheet.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: goldColor } };
  summarySheet.getCell('A2').alignment = { horizontal: 'center' };

  summarySheet.getCell('A3').font = { bold: true, size: 11 };
  summarySheet.getCell('A4').font = { italic: true, size: 10 };

  // Score Table Header
  const headerRow = summarySheet.addRow([
    'Sub-Crit',
    'Criterion Title',
    'Max Marks',
    'Awarded Marks',
    'Evaluation Summary / Metrics',
  ]);
  headerRow.font = { bold: true, color: { argb: 'FFFFFF' } };
  headerRow.fill = headerFill;

  for (const sc of report.subCriteria) {
    const row = summarySheet.addRow([
      sc.criterion,
      sc.title,
      sc.maxMarks,
      sc.awardedMarks,
      sc.summary,
    ]);
    row.getCell(3).alignment = { horizontal: 'center' };
    row.getCell(4).alignment = { horizontal: 'center' };
    row.getCell(4).font = { bold: true };
  }

  // Total Score Row
  const totalRow = summarySheet.addRow([
    'Total',
    'CRITERION 5 TOTAL SCORE',
    200,
    report.summary.totalMarks,
    `Calculated Institutional Score: ${report.summary.totalMarks} / 200`,
  ]);
  totalRow.font = { bold: true, size: 11 };
  totalRow.fill = subHeaderFill;
  totalRow.getCell(3).alignment = { horizontal: 'center' };
  totalRow.getCell(4).alignment = { horizontal: 'center' };

  // ─────────────────────────────────────────────────────────────
  // SHEET 2: TABLE 5.1 - FACULTY INFORMATION
  // ─────────────────────────────────────────────────────────────
  const fSheet = wb.addWorksheet('Table 5.1 Faculty Details', {
    views: [{ showGridLines: true }],
  });
  fSheet.columns = [
    { width: 6 },
    { width: 28 },
    { width: 22 },
    { width: 22 },
    { width: 16 },
    { width: 22 },
    { width: 20 },
  ];

  fSheet.addRow(['Table 5.1: Faculty Information Master List']);
  fSheet.mergeCells('A1:G1');
  fSheet.getCell('A1').font = { bold: true, size: 12, color: { argb: 'FFFFFF' } };
  fSheet.getCell('A1').fill = headerFill;

  const fHeader = fSheet.addRow([
    'S.No',
    'Faculty Name',
    'Designation',
    'Highest Qualification',
    'Joining Date',
    'Association Type',
    'PAN / ID Card',
  ]);
  fHeader.font = { bold: true, color: { argb: 'FFFFFF' } };
  fHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: brandColor } };

  report.tables.facultyList.forEach((f, idx) => {
    fSheet.addRow([
      idx + 1,
      f.name,
      f.designation,
      f.qualification,
      f.dateOfJoining ?? '—',
      f.natureOfAssociation,
      f.pan ?? 'Available in File',
    ]);
  });

  // ─────────────────────────────────────────────────────────────
  // SHEET 3: TABLE 5.4 - SPONSORED RESEARCH PROJECTS
  // ─────────────────────────────────────────────────────────────
  const grantSheet = wb.addWorksheet('Table 5.4 Grants', {
    views: [{ showGridLines: true }],
  });
  grantSheet.columns = [
    { width: 6 },
    { width: 36 },
    { width: 24 },
    { width: 22 },
    { width: 20 },
    { width: 18 },
    { width: 12 },
    { width: 14 },
  ];

  grantSheet.addRow(['Table 5.4: Sponsored Research Projects & Funding']);
  grantSheet.mergeCells('A1:H1');
  grantSheet.getCell('A1').font = { bold: true, size: 12, color: { argb: 'FFFFFF' } };
  grantSheet.getCell('A1').fill = headerFill;

  const gHeader = grantSheet.addRow([
    'S.No',
    'Project Title',
    'Principal Investigator',
    'Funding Agency',
    'Sanction Order No.',
    'Sanctioned (₹ Lakhs)',
    'Sanction Year',
    'Status',
  ]);
  gHeader.font = { bold: true, color: { argb: 'FFFFFF' } };
  gHeader.fill = headerFill;

  report.tables.sponsoredProjects.forEach((g, idx) => {
    grantSheet.addRow([
      idx + 1,
      g.title,
      g.pi,
      g.agency,
      g.sanctionNo,
      g.amountLakhs,
      g.year,
      g.status,
    ]);
  });

  // ─────────────────────────────────────────────────────────────
  // SHEET 4: TABLE 5.5 - CONSULTANCY PROJECTS
  // ─────────────────────────────────────────────────────────────
  const conSheet = wb.addWorksheet('Table 5.5 Consultancy', {
    views: [{ showGridLines: true }],
  });
  conSheet.columns = [
    { width: 6 },
    { width: 34 },
    { width: 24 },
    { width: 28 },
    { width: 20 },
    { width: 16 },
  ];

  conSheet.addRow(['Table 5.5: Consultancy & Corporate Training Projects']);
  conSheet.mergeCells('A1:F1');
  conSheet.getCell('A1').font = { bold: true, size: 12, color: { argb: 'FFFFFF' } };
  conSheet.getCell('A1').fill = headerFill;

  const cHeader = conSheet.addRow([
    'S.No',
    'Consultancy Title',
    'Project Coordinator',
    'Client / Organization',
    'Revenue (₹ Lakhs)',
    'Academic Year',
  ]);
  cHeader.font = { bold: true, color: { argb: 'FFFFFF' } };
  cHeader.fill = headerFill;

  report.tables.consultancyProjects.forEach((c, idx) => {
    conSheet.addRow([
      idx + 1,
      c.title,
      c.coordinator,
      c.client,
      c.revenueLakhs,
      c.academicYear,
    ]);
  });

  // ─────────────────────────────────────────────────────────────
  // SHEET 5: TABLE 5.7 - RESEARCH PUBLICATIONS & IPR
  // ─────────────────────────────────────────────────────────────
  const pubSheet = wb.addWorksheet('Table 5.7 Publications & IPR', {
    views: [{ showGridLines: true }],
  });
  pubSheet.columns = [
    { width: 6 },
    { width: 36 },
    { width: 24 },
    { width: 28 },
    { width: 14 },
    { width: 12 },
    { width: 10 },
    { width: 24 },
  ];

  pubSheet.addRow(['Table 5.7: Academic Research Publications (SCI / Scopus) & Patents']);
  pubSheet.mergeCells('A1:H1');
  pubSheet.getCell('A1').font = { bold: true, size: 12, color: { argb: 'FFFFFF' } };
  pubSheet.getCell('A1').fill = headerFill;

  const pHeader = pubSheet.addRow([
    'S.No',
    'Paper / Article Title',
    'Authors',
    'Journal Name',
    'Scopus Source ID',
    'Quartile',
    'Year',
    'DOI / Link',
  ]);
  pHeader.font = { bold: true, color: { argb: 'FFFFFF' } };
  pHeader.fill = headerFill;

  report.tables.publications.forEach((p, idx) => {
    pubSheet.addRow([
      idx + 1,
      p.title,
      p.authors,
      p.journal,
      p.scopusSourceId,
      p.quartile,
      p.year,
      p.doi,
    ]);
  });

  // ─────────────────────────────────────────────────────────────
  // SHEET 6: TABLE 5.8 - FDPS & STTPS ATTENDED
  // ─────────────────────────────────────────────────────────────
  const fdpSheet = wb.addWorksheet('Table 5.8 FDPs & STTPs', {
    views: [{ showGridLines: true }],
  });
  fdpSheet.columns = [
    { width: 6 },
    { width: 26 },
    { width: 38 },
    { width: 28 },
    { width: 16 },
    { width: 10 },
  ];

  fdpSheet.addRow(['Table 5.8: Faculty Development Programs (FDPs) & STTPs Attended']);
  fdpSheet.mergeCells('A1:F1');
  fdpSheet.getCell('A1').font = { bold: true, size: 12, color: { argb: 'FFFFFF' } };
  fdpSheet.getCell('A1').fill = headerFill;

  const fdpHeader = fdpSheet.addRow([
    'S.No',
    'Faculty Name',
    'FDP / STTP Program Title',
    'Organizing Institution',
    'Duration (Days)',
    'Year',
  ]);
  fdpHeader.font = { bold: true, color: { argb: 'FFFFFF' } };
  fdpHeader.fill = headerFill;

  report.tables.fdps.forEach((f, idx) => {
    fdpSheet.addRow([
      idx + 1,
      f.facultyName,
      f.title,
      f.organizer,
      f.durationDays,
      f.year,
    ]);
  });

  return wb;
}
