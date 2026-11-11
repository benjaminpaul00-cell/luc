import './styles.css';
document.querySelector('#app')!.innerHTML =
  '<header><div class="wrap"><h1>F10 · Facturation</h1><p>Interventions, calculs et exportation Excel</p></div></header><div class="wrap"><div class="stats"><div class="stat"><small>Total de la semaine · TPS + TVQ incluses</small><b id="sum">0,00 $</b><small id="weekRange">Semaine en cours</small></div><div class="stat"><small>Interventions cette semaine</small><b id="count">0</b></div></div><main><section class="panel"><h2 id="formTitle">Nouvelle intervention</h2><form id="invoice"><div class="grid"><div class="field"><label>Numéro de facture *</label><input name="invoiceNumber" value="F10" required placeholder="Ex. F10"></div><div class="field"><label>Tech ID</label><input class="fixed" value="x239197" disabled></div><div class="field"><label>Date du service *</label><input id="date" type="date" name="date" required></div><div class="field"><label>Numéro de job *</label><input name="job" required placeholder="Ex. 42122063"></div><div class="field wide"><label>Nom du client *</label><input name="client" required placeholder="Nom au compte"></div><div class="field"><label>Statut</label><select name="status"><option>Complete / Complet</option><option>Incomplete / Incomplet</option></select></div><div class="field"><label>Province</label><input class="fixed" value="QC" disabled></div><div class="field"><label>Point de départ (ville) *</label><input name="origin" required placeholder="Ex. Juliette"></div><div class="field"><label>Lieu de service (ville) *</label><input name="destination" required placeholder="Ex. Langelier"></div><div class="field"><label>Distance (km) facultative</label><input name="distance" placeholder="Ex. 75"></div><div class="field"><label>Temps de conduite facultatif</label><input name="driveTime" placeholder="Ex. 1.5"></div><div class="field wide"><label>Frais liés à *</label><select id="fees" name="fees"></select></div><div class="field wide"><label>Notes sur le service rendu *</label><textarea name="notes" required placeholder="Décris le travail effectué..."></textarea></div><div class="field"><label>Quantité *</label><input id="quantity" type="number" name="quantity" min="0.01" step="any" value="1" required></div><div class="field"><label>Taux fixe</label><input class="fixed" value="40 $ / unité" disabled></div></div><div class="total"><span>Charges totales</span><strong id="lineTotal">40,00 $</strong></div><button class="primary" id="submit" type="submit">Enregistrer l’intervention</button></form></section><section class="panel"><h2>Exporter les factures</h2><div class="field"><label>Choisir une date dans la période</label><input type="date" id="exportDate"></div><p class="muted">L\'export mensuel regroupe 4 semaines consécutives dans un seul classeur Excel.</p><div class="export-actions"><div class="export-pair"><button class="secondary" id="week">Exporter 1 semaine</button></div><div class="export-pair"><button class="primary" id="month">Exporter 4 semaines</button></div></div><div class="notice">Version de démonstration : données conservées uniquement sur cet appareil. L\'export utilise le fichier Excel original. Vérifie le fichier téléchargé avant de l\'envoyer au client. N\'entre pas de données confidentielles.</div></section><section class="panel"><h2>Historique des interventions</h2><div id="entries"></div></section></main></div>';
import { Buffer } from 'buffer';
import process from 'process';
(globalThis as typeof globalThis & { Buffer: typeof Buffer }).Buffer = Buffer;
(globalThis as typeof globalThis & { process: typeof process }).process = process;
import * as XLSX from 'xlsx';

const fees = [
  'Service call / Appel de Service',
  'Install / Installer',
  'Travel expenses (provide description under notes) / Dépenses de voyage (fournir une description dans les notes)',
  'Other expenses (3) / Autre Dépenses (3)',
];
const $ = s => document.querySelector(s);
const stored = () => {
  try {
    return JSON.parse(localStorage.getItem('f10-records') || '[]');
  } catch {
    return [];
  }
};
let records = stored();
let editId = null;
const money = n =>
  Number(n || 0).toLocaleString('fr-CA', {
    style: 'currency',
    currency: 'CAD',
  });
const today = new Date().toLocaleDateString('en-CA');
$('#date').value = today;
$('#fees').innerHTML = fees.map(f => '<option>' + f + '</option>').join('');
const save = () => localStorage.setItem('f10-records', JSON.stringify(records));
const dateObj = s => new Date(s + 'T12:00:00');
const monday = d => {
  const x = dateObj(d);
  x.setDate(x.getDate() - ((x.getDay() + 6) % 7));
  return x;
};
const dateKey = d =>
  d.getFullYear() +
  '-' +
  String(d.getMonth() + 1).padStart(2, '0') +
  '-' +
  String(d.getDate()).padStart(2, '0');
const weekRows = d => {
  let start = monday(d);
  let end = new Date(start);
  end.setDate(end.getDate() + 7);
  return records
    .filter(r => {
      let x = dateObj(r.date);
      return x >= start && x < end;
    })
    .sort((a, b) => a.date.localeCompare(b.date));
};
function draw() {
  const weekly = weekRows(dateKey(new Date()));
  const subtotal = weekly.reduce((a, r) => a + Number(r.quantity) * 40, 0);
  const roundCents = value => Math.round((value + Number.EPSILON) * 100) / 100;
  const tps = roundCents(subtotal * 0.05);
  const tvq = roundCents(subtotal * 0.09975);
  $('#count').textContent = weekly.length;
  $('#sum').textContent = money(subtotal + tps + tvq);
  const start = monday(dateKey(new Date()));
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  $('#weekRange').textContent = dateKey(start) + ' au ' + dateKey(end);
  $('#entries').innerHTML =
    records
      .slice()
      .sort((a, b) => b.date.localeCompare(a.date))
      .map(
        r =>
          '<div class="entry"><div><b>' +
          escapeHtml(r.client) +
          '</b><small>' +
          r.date +
          ' · Job ' +
          escapeHtml(r.job) +
          ' · ' +
          escapeHtml(r.status) +
          '</small></div><strong>' +
          money(r.quantity * 40) +
          '</strong><div class="actions"><button data-edit="' +
          r.id +
          '">Modifier</button><button data-delete="' +
          r.id +
          '">Supprimer</button></div></div>'
      )
      .join('') || '<p class="muted">Aucune intervention enregistrée.</p>';
}
function escapeHtml(v) {
  return String(v ?? '').replace(
    /[&<>"']/g,
    c =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[
        c
      ]
  );
}
$('#quantity').addEventListener(
  'input',
  () => ($('#lineTotal').textContent = money(Number($('#quantity').value) * 40))
);
$('#invoice').addEventListener('submit', e => {
  e.preventDefault();
  const data = Object.fromEntries(new FormData(e.currentTarget).entries());
  data.quantity = Number(data.quantity);
  if (
    !data.date ||
    !data.invoiceNumber?.trim() ||
    !data.job.trim() ||
    !data.client.trim() ||
    !data.origin.trim() ||
    !data.destination.trim() ||
    !data.notes.trim() ||
    !(data.quantity > 0)
  ) {
    alert(
      'Remplis tous les champs obligatoires et indique une quantité positive.'
    );
    return;
  }
  data.id = editId || String(Date.now());
  if (editId) records = records.map(r => (r.id === editId ? data : r));
  else records.push(data);
  save();
  draw();
  e.currentTarget.reset();
  $('#date').value = today;
  $('[name="invoiceNumber"]').value = 'F10';
  $('#quantity').value = '1';
  $('#lineTotal').textContent = money(40);
  editId = null;
  $('#submit').textContent = 'Enregistrer l’intervention';
  $('#formTitle').textContent = 'Nouvelle intervention';
  window.scrollTo({ top: 0, behavior: 'smooth' });
});
$('#entries').addEventListener('click', e => {
  const id = e.target.dataset.edit || e.target.dataset.delete;
  if (!id) return;
  if (e.target.dataset.delete) {
    if (confirm('Supprimer cette intervention ?')) {
      records = records.filter(r => r.id !== id);
      save();
      draw();
    }
    return;
  }
  const r = records.find(x => x.id === id);
  if (!r) return;
  editId = id;
  for (const [key, val] of Object.entries(r)) {
    const el = $('[name="' + key + '"]');
    if (el) el.value = val;
  }
  $('#lineTotal').textContent = money(r.quantity * 40);
  $('#submit').textContent = 'Enregistrer les modifications';
  $('#formTitle').textContent = 'Modifier l’intervention';
  window.scrollTo({ top: 0, behavior: 'smooth' });
});
/* Legacy generated sheet removed */
const cols = [
  'Numéro de facture',
  'Date du service/installation',
  'Numéro de job',
  'Tech XID',
  'Nom du client (Nom au compte)',
  'Complète / Incomplète',
  'Point de départ (Ville)',
  'Lieu de service (Ville)',
  'Province',
  'Distance de conduite (km)',
  'Temps de conduite estimé (hrs)',
  'Frais liés à',
  'Notes sur le service rendu ou les dépenses facturées',
  'Quantité',
  'Taux',
  'Charges Total',
];
function exportSheet(rows, title) {
  const aoa = [
    [
      'ELS ALARME ET SYSTEME DE SECURITE',
      '',
      '8080 rue Juliette',
      '',
      'LUC MARIE X239197',
      '',
      'Vendu à:',
      '',
      '',
      'Alarme Excellence Inc. 7240 Cordner, Lasalle, H8N2W8',
    ],
    [title],
    [],
    [],
    cols,
  ];
  rows.forEach(r =>
    aoa.push([
      r.invoiceNumber || 'F10',
      r.date,
      r.job,
      'x239197',
      r.client,
      r.status,
      r.origin,
      r.destination,
      'QC',
      r.distance || '',
      r.driveTime || '',
      r.fees,
      r.notes,
      Number(r.quantity),
      40,
      { f: 'N' + (aoa.length + 1) + '*O' + (aoa.length + 1) },
    ])
  );
  const n = aoa.length;
  aoa.push([
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    'TOTAL HT',
    '',
    { f: 'SUM(P6:P' + n + ')' },
  ]);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [
    { wch: 13 },
    { wch: 19 },
    { wch: 17 },
    { wch: 13 },
    { wch: 23 },
    { wch: 22 },
    { wch: 20 },
    { wch: 22 },
    { wch: 10 },
    { wch: 18 },
    { wch: 19 },
    { wch: 48 },
    { wch: 65 },
    { wch: 12 },
    { wch: 10 },
    { wch: 16 },
  ];
  return ws;
}
function legacyDownload(period) {
  const date = $('#exportDate').value || today;
  const wb = XLSX.utils.book_new();
  if (period === 'week') {
    const rows = weekRows(date);
    if (!rows.length) {
      alert('Aucune intervention pour cette semaine.');
      return;
    }
    XLSX.utils.book_append_sheet(
      wb,
      exportSheet(rows, 'Semaine du ' + dateKey(monday(date))),
      'Semaine'
    );
  } else {
    const anchor = monday(date);
    for (let i = 0; i < 4; i++) {
      const d = new Date(anchor);
      d.setDate(d.getDate() + i * 7);
      XLSX.utils.book_append_sheet(
        wb,
        exportSheet(weekRows(dateKey(d)), 'Semaine du ' + dateKey(d)),
        'Semaine ' + (i + 1)
      );
    }
  }
  XLSX.writeFile(wb, 'F10_' + period + '_' + date + '.xlsx');
}
async function download(period, share = false) {
  let date = $('#exportDate').value || today;
  if (period === 'week' && weekRows(date).length === 0 && records.length > 0) {
    const dated = records.filter(r => /^\d{4}-\d{2}-\d{2}$/.test(r.date) && !Number.isNaN(dateObj(r.date).getTime())).sort((a,b) => b.date.localeCompare(a.date));
    if (dated.length) {
      const latest = dated[0].date;
      const start = dateKey(monday(latest));
      const end = new Date(monday(latest));
      end.setDate(end.getDate() + 6);
      const ok = confirm('Aucune intervention pour la semaine sélectionnée. Une intervention existe pour la semaine du ' + start + ' au ' + dateKey(end) + '. Voulez-vous exporter cette semaine à la place ?');
      if (!ok) return;
      date = latest;
      $('#exportDate').value = latest;
    }
  }
  const weeks = [];
  const anchor = monday(date);
  for (let i = 0; i < (period === 'week' ? 1 : 4); i++) {
    const d = new Date(anchor);
    d.setDate(d.getDate() + i * 7);
    weeks.push({ date: dateKey(d), rows: weekRows(dateKey(d)) });
  }
  if (period === 'week' && weeks[0].rows.length === 0) {
    alert(records.length ? 'Aucune intervention pour cette semaine. Vérifie les dates dans l’historique.' : 'Aucune intervention enregistrée. Ajoute une intervention avant de partager.');
    return;
  }
  if (weeks.some(w => w.rows.length > 53)) {
    alert('Plus de 53 interventions dans une semaine. Export annulé pour éviter de modifier le gabarit.');
    return;
  }
  try {
    const response = await fetch('./resources/modele_original.xlsx', { cache: 'no-store' });
    if (!response.ok) throw new Error('Le gabarit Excel est indisponible.');
    const templateBytes = await response.arrayBuffer();
    // SheetJS réécrit intégralement le fichier Office Open XML au lieu de
    // conserver des éléments internes du gabarit susceptibles de poser problème.
    const template = XLSX.read(templateBytes, { type: 'array', cellStyles: true });
    const templateSheet = template.Sheets[template.SheetNames[0]];
    if (!templateSheet) throw new Error('Le gabarit ne contient aucune feuille.');
    const workbook = XLSX.utils.book_new();
    for (let i = 0; i < weeks.length; i++) {
      const sheet = JSON.parse(JSON.stringify(templateSheet));
      const { rows, date: start } = weeks[i];
      const end = new Date(dateObj(start));
      end.setDate(end.getDate() + 6);
      sheet.E3 = { t: 's', v: 'Du ' + start + ' au ' + dateKey(end) };
      for (let n = 6; n <= 58; n++) {
        for (const col of ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P']) {
          delete sheet[col + n];
        }
      }
      rows.forEach((r, index) => {
        const n = index + 6;
        const dt = dateObj(r.date);
        const serial = (Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate()) - Date.UTC(1899, 11, 30)) / 86400000;
        const values = {
          A:r.invoiceNumber || 'F10', B:serial, C:r.job, D:'x239197', E:r.client,
          F:r.status, G:r.origin, H:r.destination, I:'QC',
          J:r.distance ? Number(r.distance) : null,
          K:r.driveTime ? Number(r.driveTime) : null,
          L:r.fees, M:r.notes, N:Number(r.quantity), O:40
        };
        for (const [col, value] of Object.entries(values)) {
          if (value == null) continue;
          sheet[col + n] = typeof value === 'number' ? { t:'n', v:value } : { t:'s', v:String(value) };
        }
        sheet['B' + n].z = 'dd/mm/yyyy';
        sheet['P' + n] = { t:'n', f:'N' + n + '*O' + n };
      });
      sheet.P59 = { t:'n', f:'SUM(P6:P58)' };
      sheet.P60 = { t:'n', f:'P59*5%' };
      sheet.P61 = { t:'n', f:'P59*9.975%' };
      sheet.P62 = { t:'n', f:'SUM(P59:P61)' };
      sheet['!ref'] = 'A1:P62';
      XLSX.utils.book_append_sheet(workbook, sheet, weeks.length === 1 ? 'Détails facturation' : 'Semaine ' + (i + 1));
    }
    const output = XLSX.write(workbook, { bookType:'xlsx', type:'array', compression:true, cellStyles:true });
    const bytes = new Uint8Array(output);
    if (bytes.length < 100 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
      throw new Error('Le classeur Excel généré est invalide.');
    }
    const filename = 'F10_' + period + '_' + date + '.xlsx';
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    if (blob.size !== bytes.byteLength) {
      throw new Error('Le téléchargement Excel est incomplet.');
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    // Certains appareils Android lisent le blob après le clic.
    // Ne pas révoquer l'URL trop tôt.
    setTimeout(() => {
      URL.revokeObjectURL(url);
      a.remove();
    }, 120000);
  } catch (error) {
    console.error(error);
    alert('Impossible de générer le fichier Excel original : '+(error.message||error));
  }
}
$('#exportDate').value = today;
$('#week').addEventListener('click', () => download('week'));
$('#month').addEventListener('click', () => download('month'));
draw();
