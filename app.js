const $ = (id) => document.getElementById(id);
const num = (id) => {
  const v = parseFloat($(id).value);
  return isNaN(v) ? 0 : v;
};
const ceilTo = (v, digits=0) => {
  const m = Math.pow(10, digits);
  return Math.ceil(v * m) / m; // Excel ROUNDUP for non-negative values
};
const fmt0 = (v) => new Intl.NumberFormat(undefined, { maximumFractionDigits: 0 }).format(v);
const money0 = (v) => new Intl.NumberFormat(undefined, { style: 'currency', currency: 'AUD', maximumFractionDigits: 0 }).format(v);
const pct = (v) => isFinite(v) ? (v*100).toFixed(2) + "%" : "—";

function sanitizeInputs() {
  ['G5','G6','G9','G10','G11','G12','G29','G26'].forEach(id => {
    if ($(id).value === '') return;
    if (parseFloat($(id).value) < 0) $(id).value = 0;
  });
  ['G22','G23','G24','G25'].forEach(id => {
    if ($(id).value === '') return;
    let v = parseFloat($(id).value);
    if (isNaN(v)) v = 0;
    v = Math.max(0, Math.min(1, v));
    $(id).value = v;
  });
}

function compute() {
  sanitizeInputs();

  // Inputs
  const G5  = num('G5');
  const G6  = num('G6');
  const G9  = num('G9');
  const G10 = num('G10');
  const G11 = num('G11');
  const G12 = num('G12');
  const G29 = num('G29');

  const G26 = num('G26');
  const G25 = num('G25');
  const G24 = num('G24');
  const G23 = num('G23');
  const G22 = num('G22');

  // Derived (as per sheet)
  const G7  = G5 - G6;                       // =G5-G6
  const G8  = (G5 !== 0) ? (G7 / G5) : 0;    // =G7/G5
  const G19 = G8;                             // =G8
  const G20 = 1 - G8;                         // =100%-G8 (decimal)

  // Summary
  const J14 = G9 * 12;                        // =G9*12
  const J15 = G12;                            // =G12
  const J13 = J14 + J15;                      // =J14+J15
  const J12 = (G19 !== 0) ? (J13 * (G20 / G19)) : 0; // per sheet algebra
  const J11 = J13 + J12;                      // =J13+J12

  // Sales Required
  const J6  = Math.max(0, (G5 !== 0) ? ceilTo(J11 / G5, 0) : 0);  // =ROUNDUP(J11/G5,0)
  const K6  = J6 * G5;                       // =J6*G5

  const J7  = Math.max(0, ceilTo(J6 / 12, 0));                   // =ROUNDUP(J6/12,0)
  const K7  = J7 * G5;                       // =J7*G5

  const J8  = Math.max(0, (G10 !== 0) ? ceilTo(J6 / G10, 0) : 0);// =ROUNDUP(J6/G10,0)
  const K8  = J8 * G5;                       // =J8*G5

  const J9  = Math.max(0, (G11 !== 0) ? ceilTo(J8 / G11, 0) : 0);// =ROUNDUP(J8/G11,0)
  const K9  = J9 * G6;                       // =J9*G6 (per sheet)

  // Funnel (Rows 26..22)
  const J26 = Math.max(0, ceilTo(((G26 !== 0) ? (J6 / G26) : 0) - G29, 0)); // =ROUNDUP(J6/G26-G29,0)
  const K26 = J26 / 12;                      // =J26/12
  const L26 = (G10 !== 0) ? (J26 / G10) : 0; // =J26/$G$10
  const M26 = (G11 !== 0) ? (L26 / G11) : 0; // =L26/$G$11

  const J25 = (G25 !== 0) ? (J26 / G25) : 0; // =J26/G25
  const K25 = J25 / 12;                      // =J25/12
  const L25 = (G10 !== 0) ? (J25 / G10) : 0; // =J25/$G$10
  const M25 = (G11 !== 0) ? (L25 / G11) : 0; // =L25/$G$11

  const J24 = (G24 !== 0) ? (J25 / G24) : 0; // =J25/G24
  const K24 = ceilTo(J24 / 12, 0);           // =ROUNDUP(J24/12,0)
  const L24 = (G10 !== 0) ? ceilTo(J24 / G10, 0) : 0; // =ROUNDUP(J24/$G$10,0)
  const M24 = (G11 !== 0) ? ceilTo(L24 / G11, 0) : 0; // =ROUNDUP(L24/$G$11,0)

  const J23 = (G23 !== 0) ? (J24 / G23) : 0; // =J24/G23
  const K23 = ceilTo(J23 / 12, 0);           // =ROUNDUP(J23/12,0)
  const L23 = (G10 !== 0) ? ceilTo(J23 / G10, 0) : 0; // =ROUNDUP(J23/$G$10,0)
  const M23 = (G11 !== 0) ? ceilTo(L23 / G11, 0) : 0; // =ROUNDUP(L23/$G$11,0)

  const J22 = (G22 !== 0) ? (J23 / G22) : 0; // =J23/G22
  const K22 = ceilTo(J22 / 12, 0);           // =ROUNDUP(J22/12,0)
  const L22 = (G10 !== 0) ? ceilTo(J22 / G10, 0) : 0; // =ROUNDUP(J22/$G$10,0)
  const M22 = (G11 !== 0) ? ceilTo(L22 / G11, 0) : 0; // =ROUNDUP(L22/$G$11,0)

  const clamp0 = (x) => (x < 0 ? 0 : x);

  // Write KPI outputs
  $('G7_out').textContent  = money0(clamp0(G7));
  $('G8_out').textContent  = pct(G8);
  $('G20_out').textContent = pct(G20);

  // Summary writes
  $('J14').textContent = money0(clamp0(J14));
  $('J15').textContent = money0(clamp0(J15));
  $('J13').textContent = money0(clamp0(J13));
  $('J12').textContent = money0(clamp0(J12));
  $('J11').textContent = money0(clamp0(J11));

  // Sales required + Revenue
  $('J6').textContent = fmt0(clamp0(J6));
  $('K6').textContent = money0(clamp0(K6));

  $('J7').textContent = fmt0(clamp0(J7));
  $('K7').textContent = money0(clamp0(K7));

  $('J8').textContent = fmt0(clamp0(J8));
  $('K8').textContent = money0(clamp0(K8));

  $('J9').textContent = fmt0(clamp0(J9));
  $('K9').textContent = money0(clamp0(K9));

  // Funnel table
  $('J26').textContent = fmt0(clamp0(J26));
  $('K26').textContent = fmt0(clamp0(K26));
  $('L26').textContent = fmt0(clamp0(L26));
  $('M26').textContent = fmt0(clamp0(M26));

  $('J25').textContent = fmt0(clamp0(J25));
  $('K25').textContent = fmt0(clamp0(K25));
  $('L25').textContent = fmt0(clamp0(L25));
  $('M25').textContent = fmt0(clamp0(M25));

  $('J24').textContent = fmt0(clamp0(J24));
  $('K24').textContent = fmt0(clamp0(K24));
  $('L24').textContent = fmt0(clamp0(L24));
  $('M24').textContent = fmt0(clamp0(M24));

  $('J23').textContent = fmt0(clamp0(J23));
  $('K23').textContent = fmt0(clamp0(K23));
  $('L23').textContent = fmt0(clamp0(L23));
  $('M23').textContent = fmt0(clamp0(M23));

  $('J22').textContent = fmt0(clamp0(J22));
  $('K22').textContent = fmt0(clamp0(K22));
  $('L22').textContent = fmt0(clamp0(L22));
  $('M22').textContent = fmt0(clamp0(M22));

  $('results').style.display = 'block';
}

document.addEventListener('DOMContentLoaded', () => {
  $('calcBtn').addEventListener('click', compute);
  $('resetBtn').addEventListener('click', () => {
    document.querySelectorAll('input').forEach(el => el.value = el.defaultValue);
    $('results').style.display = 'none';
  });
});
