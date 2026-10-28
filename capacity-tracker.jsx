import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

const CSS = `
.cp {
  --bg: #f2f5f7; --panel: #ffffff; --ink: #1f2a33; --muted: #5b6b78; --line: #cbd5dc;
  --calc: #e2eeed; --accent: #0f766e; --on-accent: #ffffff;
  --c1: #3d5a80; --c2: #b7791f; --c3: #0f766e; --warn: #b42318;
  color-scheme: light dark;
  background: var(--bg); color: var(--ink);
  font: 15px/1.5 system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
  padding: 32px 20px 64px; min-height: 100vh;
}
@media (prefers-color-scheme: dark) {
  .cp {
    --bg: #12181e; --panel: #1a232b; --ink: #e7edf2; --muted: #9aa8b4; --line: #2f3d49;
    --calc: #17302f; --accent: #4fb8ae; --on-accent: #0b1a19;
    --c1: #7f9fcc; --c2: #e0a548; --c3: #4fb8ae; --warn: #f08a7e;
  }
}
.cp *, .cp *::before, .cp *::after { box-sizing: border-box; }
.cp .wrap { max-width: 960px; margin: 0 auto; }
.cp h1 { font-size: 28px; line-height: 1.2; margin: 0 0 6px; font-weight: 650; letter-spacing: -0.015em; }
.cp .sub { margin: 0 0 28px; color: var(--muted); max-width: 62ch; }
.cp .sec { border-top: 2px solid var(--ink); padding: 18px 0 30px; }
.cp h2 { font-size: 19px; margin: 0 0 4px; font-weight: 650; }
.cp .hint { margin: 0 0 16px; color: var(--muted); font-size: 14px; max-width: 70ch; }
.cp .tablewrap { overflow-x: auto; }
.cp table { width: 100%; border-collapse: collapse; }
.cp th { text-align: left; font-weight: 600; font-size: 13px; color: var(--muted); padding: 6px 8px; border-bottom: 1px solid var(--ink); }
.cp td { padding: 6px 8px; border-bottom: 1px solid var(--line); vertical-align: middle; }
.cp th.r, .cp td.r { text-align: right; }
.cp .nums { font-variant-numeric: tabular-nums; }
.cp input[type=number], .cp input[type=text], .cp select {
  font: inherit; color: var(--ink); background: var(--panel);
  border: 1px solid var(--line); border-radius: 4px; padding: 5px 8px; width: 100%; min-width: 64px;
}
.cp input[type=number] { font-variant-numeric: tabular-nums; }
.cp .capnum { text-align: right; }
.cp .calc { background: var(--calc); border-radius: 4px; padding: 5px 10px; font-weight: 650; text-align: right; font-variant-numeric: tabular-nums; }
.cp .fields { display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 16px; margin-top: 20px; align-items: start; }
.cp .fields.first { margin-top: 0; }
.cp .fields label { display: block; font-size: 13px; color: var(--muted); margin-bottom: 4px; }
.cp .formula { font-size: 13px; color: var(--muted); margin-top: 4px; font-variant-numeric: tabular-nums; }
.cp button, .cp .btn {
  font: inherit; color: var(--ink); background: var(--panel); border: 1px solid var(--line);
  border-radius: 4px; padding: 6px 14px; cursor: pointer; display: inline-block;
}
.cp button:hover, .cp .btn:hover { border-color: var(--muted); }
.cp .primary { background: var(--accent); color: var(--on-accent); border-color: var(--accent); font-weight: 600; }
.cp .sr { position: absolute; width: 1px; height: 1px; opacity: 0; overflow: hidden; }
.cp .btn:focus-within, .cp button:focus-visible, .cp input:focus-visible, .cp select:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.cp .row { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
.cp .msg { color: var(--muted); font-size: 14px; margin-top: 10px; }
.cp .msg.err { color: var(--warn); }
.cp .stats { display: flex; gap: 28px; flex-wrap: wrap; margin: 18px 0 4px; }
.cp .stat b { display: block; font-size: 26px; line-height: 1.1; font-variant-numeric: tabular-nums; }
.cp .stat span { font-size: 13px; color: var(--muted); }
.cp .colpick { margin-top: 14px; max-width: 360px; }
.cp .colpick label { display: block; font-size: 13px; color: var(--muted); margin-bottom: 4px; }
.cp .bysprint { margin-top: 18px; max-width: 520px; }
.cp .bysprint h3 { font-size: 15px; margin: 0 0 6px; }
.cp table.report { background: var(--panel); margin-top: 22px; }
.cp .report th { border: 1px solid var(--line); border-bottom: 2px solid var(--ink); padding: 8px 10px; color: var(--ink); }
.cp .report td { border: 1px solid var(--line); padding: 8px 10px; }
.cp .report td.blank { background: transparent; }
.cp .report tfoot td { background: var(--calc); }
.cp .report .big { font-size: 18px; font-weight: 650; font-variant-numeric: tabular-nums; }
.cp .goalcell { display: flex; gap: 6px; align-items: center; }
.cp .x { padding: 2px 9px; line-height: 1.2; color: var(--muted); }
.cp .add { margin-top: 12px; }
.cp .chart svg { width: 100%; height: auto; display: block; max-width: 720px; }
.cp .chartnote { margin-top: 8px; font-size: 14px; color: var(--muted); }
`;

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const fmt = (n) => String(Math.round((+n || 0) * 100) / 100);
const num = (v) => {
  const x = parseFloat(String(v).replace(",", "."));
  return isFinite(x) ? x : 0;
};
const norm = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, "");

const DEFAULT_EST = [
  { name: "Optimistic", period: "Aug \u2192 Dec 2026", sp: "15", sprints: "7" },
  { name: "Most likely", period: "Jan \u2192 Mar 2027", sp: "10", sprints: "4" },
  { name: "Pessimistic", period: "Apr \u2192 May 2027", sp: "5", sprints: "3" },
];

function computeUpload(u) {
  if (!u) return null;
  const out = { tickets: u.data.length, withSp: 0, total: 0, sprints: [] };
  if (u.sp < 0) return out;
  const map = {};
  u.data.forEach((r) => {
    const raw = r[u.sp];
    const has = String(raw).trim() !== "" && isFinite(parseFloat(String(raw).replace(",", ".")));
    const v = has ? num(raw) : 0;
    if (has) {
      out.withSp++;
      out.total += v;
    }
    let sprint = "No sprint";
    u.sprintIdx.forEach((i) => {
      const s = String(r[i]).trim();
      if (s) sprint = s;
    });
    map[sprint] = map[sprint] || { name: sprint, tickets: 0, sp: 0 };
    map[sprint].tickets++;
    map[sprint].sp += v;
  });
  out.sprints = Object.values(map).sort((a, b) =>
    a.name.localeCompare(b.name, undefined, { numeric: true })
  );
  return out;
}

function Chart({ cap, act, rem }) {
  const W = 640, rowH = 56, top = 14, labelW = 150, padR = 70;
  const leftPad = labelW + (rem < 0 ? 52 : 0);
  const H = top + rowH * 3 + 8;
  const plotW = W - leftPad - padR;
  const vmax = Math.max(cap, act, 1);
  const vmin = Math.min(0, rem);
  const X = (v) => leftPad + ((v - vmin) / (vmax - vmin)) * plotW;
  const rows = [
    ["Original capacity", cap, "var(--c1)"],
    ["Actual spent", act, "var(--c2)"],
    ["Remaining", rem, rem < 0 ? "var(--warn)" : "var(--c3)"],
  ];
  return (
    <svg
      viewBox={`0 0 ${W} ${H}`}
      role="img"
      aria-label={`Original capacity ${fmt(cap)}, actual spent ${fmt(act)}, remaining ${fmt(rem)} story points`}
    >
      <line x1={X(0)} y1={top - 4} x2={X(0)} y2={top + rowH * 3 - 6} stroke="var(--muted)" />
      {rows.map((r, i) => {
        const y = top + i * rowH + 6;
        const h = 30;
        const x0 = X(0);
        const x1 = X(r[1]);
        const pos = r[1] >= 0;
        return (
          <g key={r[0]}>
            <text x={labelW - 12} y={y + h / 2 + 5} textAnchor="end" fill="var(--ink)" fontSize="14">
              {r[0]}
            </text>
            <rect x={Math.min(x0, x1)} y={y} width={Math.max(Math.abs(x1 - x0), 1)} height={h} rx="2" fill={r[2]} />
            <text
              x={pos ? x1 + 8 : x1 - 8}
              y={y + h / 2 + 5}
              textAnchor={pos ? "start" : "end"}
              fill="var(--ink)"
              fontSize="15"
              fontWeight="600"
            >
              {fmt(r[1])}
            </text>
          </g>
        );
      })}
    </svg>
  );
}

export default function CapacityTracker() {
  const [est, setEst] = useState(DEFAULT_EST);
  const [loss, setLoss] = useState("10");
  const [focus, setFocus] = useState("0.9");
  const [goal, setGoal] = useState({ name: "1,2,3", effort: "47", cont: "7" });
  const [extra, setExtra] = useState([]);
  const [upload, setUpload] = useState(null);
  const [uploadErr, setUploadErr] = useState("");
  const nextId = useRef(1);
  const fileRef = useRef(null);

  const cap = useMemo(() => {
    const gross = est.reduce((a, e) => a + num(e.sp) * num(e.sprints), 0);
    const lossAmt = (gross * num(loss)) / 100;
    const exact = (gross - lossAmt) * num(focus);
    return { gross, lossAmt, exact, value: Math.round(exact) };
  }, [est, loss, focus]);

  const up = useMemo(() => computeUpload(upload), [upload]);

  const actual = up ? up.total : 0;
  const remaining = cap.value - actual;
  const subs = [goal, ...extra].map((g) => num(g.effort) + num(g.cont));
  const totalPlanned = subs.reduce((a, b) => a + b, 0);
  const d = new Date();
  const today = `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;

  const setEstField = (i, k, v) => setEst((list) => list.map((e, j) => (j === i ? { ...e, [k]: v } : e)));
  const setExtraField = (id, k, v) => setExtra((list) => list.map((g) => (g.id === id ? { ...g, [k]: v } : g)));

  const onFile = (e) => {
    const f = e.target.files && e.target.files[0];
    if (!f) return;
    const rd = new FileReader();
    rd.onerror = () => setUploadErr("The file could not be read. Try exporting it from Jira again.");
    rd.onload = (ev) => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        let hi = aoa.slice(0, 10).findIndex((r) => r.some((c) => norm(c).includes("storypoint")));
        if (hi < 0) hi = 0;
        const hrow = aoa[hi] || [];
        const headers = hrow.map((h, i) => String(h).trim() || `Column ${i + 1}`);
        const data = aoa.slice(hi + 1).filter((r) => r.some((c) => String(c).trim() !== ""));
        if (!data.length) {
          setUploadErr("No ticket rows found in the first sheet of this file.");
          return;
        }
        let sp = headers.findIndex((h) => ["storypoints", "storypointestimate"].includes(norm(h)));
        if (sp < 0) sp = headers.findIndex((h) => norm(h).includes("storypoint"));
        const sprintIdx = headers.map((h, i) => (norm(h).startsWith("sprint") ? i : -1)).filter((i) => i >= 0);
        setUploadErr("");
        setUpload({ name: f.name, headers, data, sp, sprintIdx });
      } catch (err) {
        setUploadErr("This file could not be parsed as Excel or CSV: " + err.message);
      }
    };
    rd.readAsArrayBuffer(f);
  };

  const clearUpload = () => {
    setUpload(null);
    setUploadErr("");
    if (fileRef.current) fileRef.current.value = "";
  };

  let upMsg = "No file uploaded yet. Actual spent effort counts as 0 until you upload one.";
  let upIsErr = false;
  if (uploadErr) {
    upMsg = uploadErr;
    upIsErr = true;
  } else if (upload && upload.sp < 0) {
    upMsg = "No Story Points column was found. Select the column that holds Story Points.";
    upIsErr = true;
  } else if (upload) {
    upMsg = `Story Points are read from the column "${upload.headers[upload.sp]}".`;
  }

  const showStats = up && upload.sp >= 0;

  return (
    <div className="cp">
      <style>{CSS}</style>
      <div className="wrap">
        <h1>Delivery capacity and burn tracker</h1>
        <p className="sub">
          Set the annual capacity, import the Jira export to get actual effort, then compare the two.
        </p>

        <section className="sec">
          <h2>Annual delivery capacity</h2>
          <p className="hint">
            Enter story points per sprint and the number of sprints for each scenario. Planned annual delivery
            capacity is calculated from the totals.
          </p>
          <div className="tablewrap">
            <table>
              <thead>
                <tr>
                  <th>Estimate</th>
                  <th>Period</th>
                  <th className="r">Story points per sprint</th>
                  <th className="r">Sprints</th>
                  <th className="r">Total (SP)</th>
                </tr>
              </thead>
              <tbody>
                {est.map((e, i) => (
                  <tr key={e.name}>
                    <td>
                      <b>{e.name}</b>
                    </td>
                    <td>
                      <input
                        type="text"
                        value={e.period}
                        aria-label={`${e.name} period`}
                        onChange={(ev) => setEstField(i, "period", ev.target.value)}
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        className="capnum"
                        min="0"
                        step="any"
                        value={e.sp}
                        aria-label={`${e.name} story points per sprint`}
                        onChange={(ev) => setEstField(i, "sp", ev.target.value)}
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        className="capnum"
                        min="0"
                        step="1"
                        value={e.sprints}
                        aria-label={`${e.name} sprint count`}
                        onChange={(ev) => setEstField(i, "sprints", ev.target.value)}
                      />
                    </td>
                    <td>
                      <div className="calc">{fmt(num(e.sp) * num(e.sprints))}</div>
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan={4}>
                    <b>Gross team capacity</b>
                  </td>
                  <td>
                    <div className="calc">{fmt(cap.gross)}</div>
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
          <div className="fields">
            <div>
              <label htmlFor="loss">Expected availability loss (%)</label>
              <input
                type="number"
                id="loss"
                min="0"
                max="100"
                step="any"
                value={loss}
                onChange={(e) => setLoss(e.target.value)}
              />
              <div className="formula">Holidays, vacation, business stuck</div>
            </div>
            <div>
              <label htmlFor="focus">Focus factor</label>
              <input
                type="number"
                id="focus"
                min="0"
                max="1"
                step="0.01"
                value={focus}
                onChange={(e) => setFocus(e.target.value)}
              />
              <div className="formula">Context switching. Use 0.85 to 0.9</div>
            </div>
            <div>
              <label>Planned annual delivery capacity</label>
              <div className="calc">{fmt(cap.value)} SP</div>
              <div className="formula">
                ({fmt(cap.gross)} &minus; {fmt(cap.lossAmt)}) &times; {fmt(num(focus))} = {fmt(cap.exact)}
                {Math.abs(cap.exact - cap.value) > 1e-9 ? `, rounded to ${fmt(cap.value)}` : ""}
              </div>
            </div>
          </div>
        </section>

        <section className="sec">
          <h2>Actual effort from Jira</h2>
          <p className="hint">
            Upload the Excel file exported from Jira. The tool adds up the Story Points column for every ticket in
            the file.
          </p>
          <div className="row">
            <label className="btn primary">
              <input ref={fileRef} type="file" className="sr" accept=".xlsx,.xls,.csv" onChange={onFile} />
              Upload Jira export
            </label>
            {upload && (
              <button type="button" onClick={clearUpload}>
                Remove file
              </button>
            )}
            {upload && (
              <span className="msg" style={{ margin: 0 }}>
                {upload.name}
              </span>
            )}
          </div>
          <div className={"msg" + (upIsErr ? " err" : "")}>{upMsg}</div>

          {upload && (
            <div className="colpick">
              <label htmlFor="colSel">Story Points column</label>
              <select
                id="colSel"
                value={String(upload.sp)}
                onChange={(e) => setUpload({ ...upload, sp: parseInt(e.target.value, 10) })}
              >
                {upload.sp < 0 && <option value="-1">Select a column&hellip;</option>}
                {upload.headers.map((h, i) => (
                  <option key={i} value={String(i)}>
                    {h}
                  </option>
                ))}
              </select>
            </div>
          )}

          {showStats && (
            <div className="stats">
              <div className="stat">
                <b>{up.tickets}</b>
                <span>tickets in file</span>
              </div>
              <div className="stat">
                <b>{up.withSp}</b>
                <span>tickets with Story Points</span>
              </div>
              <div className="stat">
                <b>{fmt(up.total)}</b>
                <span>total Story Points</span>
              </div>
            </div>
          )}

          {showStats && up.sprints.length > 0 && (
            <div className="bysprint">
              <h3>Story Points by sprint</h3>
              <table>
                <thead>
                  <tr>
                    <th>Sprint</th>
                    <th className="r">Tickets</th>
                    <th className="r">Story Points</th>
                  </tr>
                </thead>
                <tbody>
                  {up.sprints.map((s) => (
                    <tr key={s.name}>
                      <td>{s.name}</td>
                      <td className="r nums">{s.tickets}</td>
                      <td className="r nums">{fmt(s.sp)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div className="msg">Each ticket counts toward the last Sprint column that has a value.</div>
            </div>
          )}
        </section>

        <section className="sec">
          <h2>Planning report</h2>
          <p className="hint">The first row comes from the three fields below. Add rows for other goals.</p>
          <div className="fields first">
            <div>
              <label htmlFor="gName">Goals</label>
              <input
                type="text"
                id="gName"
                value={goal.name}
                onChange={(e) => setGoal({ ...goal, name: e.target.value })}
              />
            </div>
            <div>
              <label htmlFor="gEffort">Planned effort committed goals</label>
              <input
                type="number"
                id="gEffort"
                min="0"
                step="any"
                value={goal.effort}
                onChange={(e) => setGoal({ ...goal, effort: e.target.value })}
              />
            </div>
            <div>
              <label htmlFor="gCont">Planned contingency for uncertainty</label>
              <input
                type="number"
                id="gCont"
                min="0"
                step="any"
                value={goal.cont}
                onChange={(e) => setGoal({ ...goal, cont: e.target.value })}
              />
            </div>
          </div>

          <div className="tablewrap">
            <table className="report">
              <thead>
                <tr>
                  <th>Goals</th>
                  <th>Planned effort committed goals</th>
                  <th>Planned contingency for uncertainty</th>
                  <th>Planned annual delivery capacity</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>{goal.name || "-"}</td>
                  <td className="nums">{fmt(num(goal.effort))}</td>
                  <td className="nums">{fmt(num(goal.cont))}</td>
                  <td className="blank"></td>
                </tr>
                {extra.map((g) => (
                  <tr key={g.id}>
                    <td>
                      <div className="goalcell">
                        <input
                          type="text"
                          aria-label="Goal name"
                          value={g.name}
                          onChange={(e) => setExtraField(g.id, "name", e.target.value)}
                        />
                        <button
                          type="button"
                          className="x"
                          title="Remove row"
                          aria-label="Remove row"
                          onClick={() => setExtra((list) => list.filter((x) => x.id !== g.id))}
                        >
                          &times;
                        </button>
                      </div>
                    </td>
                    <td>
                      <input
                        type="number"
                        min="0"
                        step="any"
                        aria-label="Planned effort committed goals"
                        value={g.effort}
                        onChange={(e) => setExtraField(g.id, "effort", e.target.value)}
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        min="0"
                        step="any"
                        aria-label="Planned contingency for uncertainty"
                        value={g.cont}
                        onChange={(e) => setExtraField(g.id, "cont", e.target.value)}
                      />
                    </td>
                    <td className="blank"></td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan={3}>
                    Total planned effort committed goals + contingency for uncertainty
                    <br />
                    <span className="big nums">
                      {subs.length === 1
                        ? `${fmt(num(goal.effort))} + ${fmt(num(goal.cont))} = ${fmt(totalPlanned)}`
                        : `${subs.map(fmt).join(" + ")} = ${fmt(totalPlanned)}`}
                    </span>
                  </td>
                  <td>
                    Original: <span className="big nums">{fmt(cap.value)}</span>
                  </td>
                </tr>
                <tr>
                  <td colSpan={3}>
                    Actual spent effort: <span className="big nums">{fmt(actual)}</span>
                  </td>
                  <td>
                    Remaining on {today}: <span className="big nums">{fmt(remaining)}</span>
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
          <button
            type="button"
            className="add"
            onClick={() => setExtra((list) => [...list, { id: nextId.current++, name: "New goal", effort: "0", cont: "0" }])}
          >
            Add goal row
          </button>
        </section>

        <section className="sec">
          <h2>Capacity and actual spend</h2>
          <p className="hint">
            All three bars use the same scale. Remaining is the original capacity minus actual spent effort.
          </p>
          <div className="chart">
            <Chart cap={cap.value} act={actual} rem={remaining} />
          </div>
          <div className="chartnote">
            {cap.value > 0
              ? `Spent ${fmt((actual / cap.value) * 100)}% of the original capacity. ` +
                (remaining < 0 ? `Over capacity by ${fmt(-remaining)} SP.` : `${fmt((remaining / cap.value) * 100)}% remains.`)
              : "Enter capacity values to see the comparison."}
          </div>
        </section>
      </div>
    </div>
  );
}
