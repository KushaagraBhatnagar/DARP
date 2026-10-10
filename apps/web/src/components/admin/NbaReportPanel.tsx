'use client';

import { useState, useEffect, useCallback } from 'react';
import { Button, Card, CardHeader, Notice, StatePill, Table, Th, Td } from '@/components/ui';
import type { DepartmentRow } from './types';

interface NbaSubCriteria {
  criterion: string;
  title: string;
  maxMarks: number;
  awardedMarks: number;
  summary: string;
  details: Record<string, unknown>;
}

interface NbaReportData {
  department: {
    id: string;
    code: string;
    name: string;
  };
  cycle: {
    id: string;
    name: string;
    startYear: number;
    endYear: number;
  };
  summary: {
    totalMarks: number;
    maxMarks: number;
    facultyCount: number;
    phdCount: number;
    sfr: number;
    fqIndex: number;
    retentionRate: number;
    totalGrantsLakhs: number;
    totalConsultancyLakhs: number;
    scopusPubsCount: number;
    patentsCount: number;
  };
  subCriteria: NbaSubCriteria[];
  tables: {
    facultyList: Array<{
      id: string;
      name: string;
      designation: string;
      qualification: string;
      dateOfJoining: string | null;
      pan: string | null;
      natureOfAssociation: string;
    }>;
    sponsoredProjects: Array<{
      id: string;
      title: string;
      pi: string;
      agency: string;
      sanctionNo: string;
      amountLakhs: number;
      year: number;
      status: string;
    }>;
    consultancyProjects: Array<{
      id: string;
      title: string;
      coordinator: string;
      client: string;
      revenueLakhs: number;
      academicYear: string;
    }>;
    publications: Array<{
      id: string;
      title: string;
      authors: string;
      journal: string;
      scopusSourceId: string;
      quartile: string;
      year: number;
      doi: string;
    }>;
    patents: Array<{
      id: string;
      title: string;
      inventors: string;
      status: string;
      patentNo: string;
      filingDate: string;
      grantDate: string;
    }>;
    fdps: Array<{
      id: string;
      facultyName: string;
      title: string;
      organizer: string;
      durationDays: number;
      year: number;
    }>;
  };
}

export function NbaReportPanel() {
  const [departmentsList, setDepartmentsList] = useState<DepartmentRow[]>([]);
  const [selectedDeptId, setSelectedDeptId] = useState<string>('');
  const [report, setReport] = useState<NbaReportData | null>(null);
  const [loading, setLoading] = useState(true);
  const [downloading, setDownloading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [activeTableTab, setActiveTableTab] = useState<'faculty' | 'grants' | 'consultancy' | 'pubs' | 'fdps'>('faculty');

  // Load departments
  useEffect(() => {
    async function loadDepts() {
      try {
        const res = await fetch('/api/admin/users');
        const json = await res.json();
        if (json.ok && json.data.departments) {
          setDepartmentsList(json.data.departments);
          if (json.data.departments.length > 0) {
            setSelectedDeptId(json.data.departments[0].id);
          }
        }
      } catch {
        setError('Could not load departments list.');
      }
    }
    loadDepts();
  }, []);

  // Fetch report data
  const fetchReport = useCallback(async (deptId: string) => {
    if (!deptId) return;
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`/api/admin/nba?departmentId=${deptId}`);
      const json = await res.json();
      if (json.ok) {
        setReport(json.data);
      } else {
        setError(json.error.message);
      }
    } catch {
      setError('Failed to fetch NBA compliance report.');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    if (selectedDeptId) {
      fetchReport(selectedDeptId);
    }
  }, [selectedDeptId, fetchReport]);

  const handleDownloadExcel = async () => {
    if (!selectedDeptId) return;
    setDownloading(true);
    try {
      const link = document.createElement('a');
      link.href = `/api/admin/nba?departmentId=${selectedDeptId}&format=xlsx`;
      link.setAttribute('download', `NBA_Tier1_SAR_${report?.department.code ?? 'Dept'}.xlsx`);
      document.body.appendChild(link);
      link.click();
      link.remove();
    } catch {
      setError('Excel download failed.');
    } finally {
      setTimeout(() => setDownloading(false), 2000);
    }
  };

  return (
    <div className="space-y-6">
      {/* Top Filter Bar */}
      <Card>
        <div className="flex flex-wrap items-center justify-between gap-4 p-4">
          <div className="flex flex-wrap items-center gap-3">
            <label htmlFor="dept-select" className="text-xs font-semibold text-ink-muted">
              Select Department:
            </label>
            <select
              id="dept-select"
              value={selectedDeptId}
              onChange={(e) => setSelectedDeptId(e.target.value)}
              className="rounded-md border border-line bg-surface px-3 py-1.5 text-sm font-medium shadow-xs focus:border-brand-500 focus:outline-hidden"
            >
              {departmentsList.map((d) => (
                <option key={d.id} value={d.id}>
                  {d.name} ({d.code})
                </option>
              ))}
            </select>
          </div>

          <div className="flex items-center gap-3">
            <Button
              variant="primary"
              onClick={handleDownloadExcel}
              disabled={downloading || !report}
            >
              {downloading ? 'Generating Excel...' : '📊 Download NBA SAR Excel (.xlsx)'}
            </Button>
          </div>
        </div>
      </Card>

      {error ? <Notice tone="danger">{error}</Notice> : null}

      {loading ? (
        <div className="flex min-h-[240px] items-center justify-center rounded-lg border border-line-faint bg-surface p-8">
          <div className="text-center">
            <div className="inline-block h-6 w-6 animate-spin rounded-full border-2 border-brand-600 border-t-transparent" />
            <p className="mt-2 text-xs text-ink-muted">Computing NBA Tier-1 SAR Criterion 5 metrics...</p>
          </div>
        </div>
      ) : report ? (
        <>
          {/* Executive Scorecard & KPIs */}
          <div className="grid grid-cols-2 gap-4 sm:grid-cols-3 lg:grid-cols-6">
            <div className="rounded-lg border border-brand-200 bg-brand-50/50 p-4 shadow-2xs">
              <div className="text-2xs font-bold uppercase tracking-wider text-brand-700">Criterion 5 Score</div>
              <div className="mt-1 text-2xl font-black text-brand-900">
                {report.summary.totalMarks} <span className="text-xs font-normal text-ink-muted">/ 200</span>
              </div>
              <div className="mt-1 text-2xs text-brand-700">
                {Math.round((report.summary.totalMarks / 200) * 100)}% Achievement
              </div>
            </div>

            <div className="rounded-lg border border-line bg-surface p-4 shadow-2xs">
              <div className="text-2xs font-bold uppercase tracking-wider text-ink-muted">Student-Faculty (SFR)</div>
              <div className="mt-1 text-2xl font-black text-ink">1:{report.summary.sfr}</div>
              <div className="mt-1 text-2xs text-success-600">AICTE Target &le; 1:15</div>
            </div>

            <div className="rounded-lg border border-line bg-surface p-4 shadow-2xs">
              <div className="text-2xs font-bold uppercase tracking-wider text-ink-muted">Ph.D. Faculty (FQ)</div>
              <div className="mt-1 text-2xl font-black text-ink">
                {report.summary.phdCount} <span className="text-xs font-normal text-ink-muted">/ {report.summary.facultyCount}</span>
              </div>
              <div className="mt-1 text-2xs text-info-600">FQ Score: {report.summary.fqIndex} / 20</div>
            </div>

            <div className="rounded-lg border border-line bg-surface p-4 shadow-2xs">
              <div className="text-2xs font-bold uppercase tracking-wider text-ink-muted">Faculty Retention</div>
              <div className="mt-1 text-2xl font-black text-ink">{report.summary.retentionRate}%</div>
              <div className="mt-1 text-2xs text-success-600">&ge; 90% (Full 10 Marks)</div>
            </div>

            <div className="rounded-lg border border-line bg-surface p-4 shadow-2xs">
              <div className="text-2xs font-bold uppercase tracking-wider text-ink-muted">Sponsored Grants</div>
              <div className="mt-1 text-2xl font-black text-ink">₹ {report.summary.totalGrantsLakhs.toFixed(1)}L</div>
              <div className="mt-1 text-2xs text-brand-600">30 / 30 Marks</div>
            </div>

            <div className="rounded-lg border border-line bg-surface p-4 shadow-2xs">
              <div className="text-2xs font-bold uppercase tracking-wider text-ink-muted">Consultancy Revenue</div>
              <div className="mt-1 text-2xl font-black text-ink">₹ {report.summary.totalConsultancyLakhs.toFixed(1)}L</div>
              <div className="mt-1 text-2xs text-brand-600">20 / 20 Marks</div>
            </div>
          </div>

          {/* Sub-Criteria Breakdown Table */}
          <Card>
            <CardHeader
              title="Criterion 5: Marks & Evaluation Breakdown"
              subtitle="Calculated marks per NBA Tier-1 Undergraduate Engineering accreditation guidelines."
            />
            <div className="overflow-x-auto">
              <Table>
                <thead>
                  <tr>
                    <Th className="w-16">Crit.</Th>
                    <Th>Sub-Criterion Title</Th>
                    <Th className="w-24 text-center">Max</Th>
                    <Th className="w-24 text-center">Awarded</Th>
                    <Th>Assessment Findings & Metrics</Th>
                  </tr>
                </thead>
                <tbody>
                  {report.subCriteria.map((sc) => (
                    <tr key={sc.criterion} className="hover:bg-surface-inset/40">
                      <Td className="font-mono text-xs font-bold text-brand-700">{sc.criterion}</Td>
                      <Td className="font-semibold text-ink">{sc.title}</Td>
                      <Td className="text-center font-mono text-xs text-ink-muted">{sc.maxMarks}</Td>
                      <Td className="text-center font-mono text-sm font-bold text-ink">
                        <span className={sc.awardedMarks === sc.maxMarks ? 'text-success-700' : 'text-brand-800'}>
                          {sc.awardedMarks}
                        </span>
                      </Td>
                      <Td className="text-xs text-ink-muted">{sc.summary}</Td>
                    </tr>
                  ))}
                  <tr className="bg-surface-inset font-bold">
                    <td colSpan={2} className="px-3 py-2 text-sm font-bold text-ink">
                      TOTAL CRITERION 5 SCORE
                    </td>
                    <Td className="text-center font-mono text-sm">200</Td>
                    <Td className="text-center font-mono text-sm font-black text-brand-900">
                      {report.summary.totalMarks}
                    </Td>
                    <Td className="text-xs text-ink-muted">
                      Overall Compliance Rate: {Math.round((report.summary.totalMarks / 200) * 100)}%
                    </Td>
                  </tr>
                </tbody>
              </Table>
            </div>
          </Card>

          {/* Interactive NBA Tables Section */}
          <Card>
            <div className="border-b border-line px-4 pt-3">
              <div className="flex flex-wrap gap-2">
                {[
                  { key: 'faculty', label: `Table 5.1: Faculty Details (${report.tables.facultyList.length})` },
                  { key: 'grants', label: `Table 5.4: Grants (${report.tables.sponsoredProjects.length})` },
                  { key: 'consultancy', label: `Table 5.5: Consultancy (${report.tables.consultancyProjects.length})` },
                  { key: 'pubs', label: `Table 5.7: Publications & IPR (${report.tables.publications.length})` },
                  { key: 'fdps', label: `Table 5.8: FDPs Attended (${report.tables.fdps.length})` },
                ].map((t) => (
                  <button
                    key={t.key}
                    type="button"
                    onClick={() => setActiveTableTab(t.key as typeof activeTableTab)}
                    className={`border-b-2 px-3 py-2 text-xs font-semibold transition-colors ${
                      activeTableTab === t.key
                        ? 'border-brand-600 text-brand-700'
                        : 'border-transparent text-ink-muted hover:text-ink'
                    }`}
                  >
                    {t.label}
                  </button>
                ))}
              </div>
            </div>

            <div className="p-4">
              {activeTableTab === 'faculty' ? (
                <Table>
                  <thead>
                    <tr>
                      <Th>Faculty Name</Th>
                      <Th>Designation</Th>
                      <Th>Highest Qualification</Th>
                      <Th>Joining Date</Th>
                      <Th>Association</Th>
                    </tr>
                  </thead>
                  <tbody>
                    {report.tables.facultyList.map((f) => (
                      <tr key={f.id}>
                        <Td className="font-semibold text-ink">{f.name}</Td>
                        <Td>{f.designation}</Td>
                        <Td>
                          <span className={`inline-flex rounded-xs px-1.5 py-0.5 text-2xs font-semibold ${
                            f.qualification.includes('Ph.D') ? 'bg-success-50 text-success-700 border border-success-200' : 'bg-surface-inset text-ink-muted'
                          }`}>
                            {f.qualification}
                          </span>
                        </Td>
                        <Td className="font-mono text-xs">{f.dateOfJoining ?? '—'}</Td>
                        <Td>{f.natureOfAssociation}</Td>
                      </tr>
                    ))}
                  </tbody>
                </Table>
              ) : null}

              {activeTableTab === 'grants' ? (
                <Table>
                  <thead>
                    <tr>
                      <Th>Project Title</Th>
                      <Th>Principal Investigator</Th>
                      <Th>Funding Agency</Th>
                      <Th>Sanction No.</Th>
                      <Th className="text-right">Amount (₹ Lakhs)</Th>
                      <Th>Status</Th>
                    </tr>
                  </thead>
                  <tbody>
                    {report.tables.sponsoredProjects.map((p) => (
                      <tr key={p.id}>
                        <Td className="font-semibold text-ink">{p.title}</Td>
                        <Td>{p.pi}</Td>
                        <Td>{p.agency}</Td>
                        <Td className="font-mono text-xs">{p.sanctionNo}</Td>
                        <Td className="text-right font-mono font-bold text-brand-800">₹ {p.amountLakhs.toFixed(2)}L</Td>
                        <Td>
                          <StatePill status="approved" />
                        </Td>
                      </tr>
                    ))}
                  </tbody>
                </Table>
              ) : null}

              {activeTableTab === 'consultancy' ? (
                <Table>
                  <thead>
                    <tr>
                      <Th>Consultancy Title</Th>
                      <Th>Coordinator</Th>
                      <Th>Client Organization</Th>
                      <Th className="text-right">Revenue (₹ Lakhs)</Th>
                      <Th>Academic Year</Th>
                    </tr>
                  </thead>
                  <tbody>
                    {report.tables.consultancyProjects.map((c) => (
                      <tr key={c.id}>
                        <Td className="font-semibold text-ink">{c.title}</Td>
                        <Td>{c.coordinator}</Td>
                        <Td>{c.client}</Td>
                        <Td className="text-right font-mono font-bold text-success-800">₹ {c.revenueLakhs.toFixed(2)}L</Td>
                        <Td className="font-mono text-xs">{c.academicYear}</Td>
                      </tr>
                    ))}
                  </tbody>
                </Table>
              ) : null}

              {activeTableTab === 'pubs' ? (
                <Table>
                  <thead>
                    <tr>
                      <Th>Paper / Article Title</Th>
                      <Th>Authors</Th>
                      <Th>Journal Name</Th>
                      <Th>Quartile</Th>
                      <Th>Year</Th>
                    </tr>
                  </thead>
                  <tbody>
                    {report.tables.publications.map((p) => (
                      <tr key={p.id}>
                        <Td className="font-semibold text-ink">{p.title}</Td>
                        <Td className="text-xs text-ink-muted">{p.authors}</Td>
                        <Td>{p.journal}</Td>
                        <Td>
                          <span className={`inline-flex rounded-xs px-1.5 py-0.5 text-2xs font-bold ${
                            p.quartile === 'Q1' ? 'bg-amber-100 text-amber-900 border border-amber-300' : 'bg-slate-100 text-slate-800'
                          }`}>
                            {p.quartile}
                          </span>
                        </Td>
                        <Td className="font-mono text-xs">{p.year}</Td>
                      </tr>
                    ))}
                  </tbody>
                </Table>
              ) : null}

              {activeTableTab === 'fdps' ? (
                <Table>
                  <thead>
                    <tr>
                      <Th>Faculty Name</Th>
                      <Th>Program Title</Th>
                      <Th>Organizing Body</Th>
                      <Th>Duration</Th>
                      <Th>Year</Th>
                    </tr>
                  </thead>
                  <tbody>
                    {report.tables.fdps.map((f) => (
                      <tr key={f.id}>
                        <Td className="font-semibold text-ink">{f.facultyName}</Td>
                        <Td>{f.title}</Td>
                        <Td>{f.organizer}</Td>
                        <Td className="font-mono text-xs">{f.durationDays} Days</Td>
                        <Td className="font-mono text-xs">{f.year}</Td>
                      </tr>
                    ))}
                  </tbody>
                </Table>
              ) : null}
            </div>
          </Card>
        </>
      ) : null}
    </div>
  );
}
