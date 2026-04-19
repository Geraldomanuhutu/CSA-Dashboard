import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import { PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend } from "recharts";
import * as XLSX from "xlsx";

const C = {
  effective:"#22c55e",ineffective:"#ef4444",mekaar:"#3b82f6",ulamm:"#8b5cf6",frp:"#ec4899",
  bg:"#0a0e1a",card:"#111827",border:"#1e293b",text:"#e2e8f0",muted:"#94a3b8",accent:"#38bdf8",
};
const STS={overdue:{label:"OVERDUE",color:"#ef4444",bg:"#ef444422",p:0},critical:{label:"CRITICAL",color:"#f97316",bg:"#f9731622",p:1},warning:{label:"WARNING",color:"#eab308",bg:"#eab30822",p:2},ontrack:{label:"ON TRACK",color:"#22c55e",bg:"#22c55e22",p:3}};
const TSTS={selesai:{label:"Selesai",color:"#22c55e",bg:"#22c55e18",icon:"✅"},proses:{label:"Dalam Proses",color:"#38bdf8",bg:"#38bdf818",icon:"🔄"},belum:{label:"Belum Dimulai",color:"#94a3b8",bg:"#94a3b818",icon:"⏳"},terkendala:{label:"Terkendala",color:"#f97316",bg:"#f9731618",icon:"⚠️"}};
const TODAY=new Date();
function getDL(t){const y=parseInt((t.match(/\d{4}/)||["2026"])[0]);if(t.startsWith("Q1"))return new Date(y,2,31);if(t.startsWith("Q2"))return new Date(y,5,30);if(t.startsWith("Q3"))return new Date(y,8,30);if(t.startsWith("Q4"))return new Date(y,11,31);if(t.toLowerCase().includes("mei"))return new Date(y,4,31);return new Date(y,11,31);}
function getSts(t){const d=getDL(t),diff=Math.ceil((d-TODAY)/864e5);if(diff<0)return{...STS.overdue,days:Math.abs(diff),diff,deadline:d,msg:`${Math.abs(diff)} hari melewati deadline`};if(diff<=30)return{...STS.critical,days:diff,diff,deadline:d,msg:`${diff} hari menuju deadline`};if(diff<=90)return{...STS.warning,days:diff,diff,deadline:d,msg:`${diff} hari menuju deadline`};return{...STS.ontrack,days:diff,diff,deadline:d,msg:`${diff} hari menuju deadline`};}
function fmtD(d){return d.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"});}
const cc=cy=>cy==="Mekaar"?C.mekaar:cy==="ULaMM"?C.ulamm:C.frp;

function parseTanggapanText(text) {
  if (!text || !String(text).trim()) return [];
  return String(text).split(/\n/).map(line => line.trim()).filter(Boolean);
}

function parseXlsx(ab) {
  const wb = XLSX.read(ab, { type: "array" });
  const raw = (name) => { const ws = wb.Sheets[name]; return ws ? XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false }) : []; };
  const cr = raw("Config"), cfg = {};
  for (let i = 3; i < cr.length; i++) { const [f, v] = cr[i]; if (f) cfg[String(f).trim()] = v; }
  const cyR = raw("Cycles"), cyH = cyR[2] || [], cycles = [];
  for (let i = 3; i < cyR.length; i++) { const r = cyR[i]; if (!r[0]) continue; const o = {}; cyH.forEach((h, ci) => { o[String(h).trim()] = r[ci] }); cycles.push({ name: String(o.name || ""), total: +o.total || 0, tested: +o.tested || 0, effective: +o.effective || 0, ineffective: +o.ineffective || 0, manualE: +o.manualE || 0, manualI: +o.manualI || 0, otomatisE: +o.otomatisE || 0, otomatisI: +o.otomatisI || 0 }); }
  const rcR = raw("RootCauses"), rcH = rcR[2] || [], rootCauses = [];
  for (let i = 3; i < rcR.length; i++) { const r = rcR[i]; if (!r[0]) continue; const o = {}; rcH.forEach((h, ci) => { o[String(h).trim()] = r[ci] }); rootCauses.push({ name: String(o.name), count: +o.count || 0, color: String(o.color || "#94a3b8") }); }
  const fR = raw("Findings"), fH = fR[2] || [], findings = [];
  for (let i = 3; i < fR.length; i++) {
    const r = fR[i]; if (!r[0]) continue; const o = {}; fH.forEach((h, ci) => { o[String(h).trim()] = r[ci] });
    const rem = [];
    for (let j = 1; j <= 10; j++) { const v = o[`remediasi_${j}`]; if (v && String(v).trim()) rem.push(String(v).trim()); }
    findings.push({
      ref: String(o.ref || ""), cycle: String(o.cycle || ""), type: String(o.type || ""),
      cause: String(o.cause || ""), desc: String(o.desc || ""), target: String(o.target || ""),
      gap: String(o.gap || ""), pic: String(o.pic || ""), remediasi: rem,
      tanggapan_status: String(o.tanggapan_status || "belum").trim().toLowerCase(),
      tanggapan_tgl: String(o.tanggapan_tgl || "").trim(),
      tanggapan_entries: parseTanggapanText(o.tanggapan),
    });
  }
  return { id: String(cfg.period_id || "uploaded"), data: { label: String(cfg.period_label || "Uploaded"), subtitle: String(cfg.period_subtitle || "Data dari Excel"), totalControls: +cfg.total_controls || 0, totalTested: +cfg.total_tested || 0, keyNotTested: +cfg.key_not_tested || 0, totalEffective: +cfg.total_effective || 0, totalIneffective: +cfg.total_ineffective || 0, manualTotal: { efektif: +cfg.manual_efektif || 0, inefektif: +cfg.manual_inefektif || 0 }, otomatisTotal: { efektif: +cfg.otomatis_efektif || 0, inefektif: +cfg.otomatis_inefektif || 0 }, cycles, rootCauses, findings, scope: { periode: String(cfg.scope_periode || ""), lokasi: String(cfg.scope_lokasi || ""), siklus: String(cfg.scope_siklus || ""), divisi: String(cfg.scope_divisi || "") } } };
}

const FINDINGS=[
{ref:"ITDM-MKR-01",cycle:"Mekaar",type:"Manual (ITDM)",cause:"Ketidaklengkapan Data",desc:"Pejabat berwenang melakukan reviu dan persetujuan atas tarikan FP4 pada Mekaar Digi — tanda tangan FP4 tidak tervalidasi dan tidak terbaca",target:"Q1 2026",gap:"GAP-MKR-05",pic:"Divisi ATI",remediasi:["Mengembangkan fitur penguncian otomatis pada dokumen apabila seluruh pihak belum menyelesaikan penandatanganan","Memastikan verifikasi tanda tangan FP4 melalui face recognition sesuai ketentuan, termasuk melibatkan legal dalam UAT","Menyesuaikan sistem membedakan FP4 dengan/tanpa foto domisili"],tanggapan_status:"proses",tanggapan_tgl:"28 Mar 2026",tanggapan_entries:["Fitur penguncian otomatis sedang dalam tahap pengembangan oleh tim ATI, target UAT minggu ke-3 April 2026","Koordinasi dengan fungsi legal sudah dilakukan untuk memastikan persyaratan PSE terpenuhi dalam UAT","Penyesuaian sistem untuk membedakan FP4 menunggu finalisasi fitur penguncian terlebih dahulu"]},
{ref:"MKR-71-02",cycle:"Mekaar",type:"Manual",cause:"Ketidakakuratan Informasi",desc:"Kepala Divisi AMK melakukan reviu atas hasil perhitungan bunga EIR bulanan — perhitungan tidak mempertimbangkan waktu pembayaran aktual dan bukti peninjauan belum terdokumentasi",target:"Q4 2026",gap:"GAP-MKR-10 / GAP-MKR-11",pic:"Divisi AMK & MRE",remediasi:["Mengintegrasikan data realisasi pembayaran ke perhitungan akrual bunga","Divisi AMK dan MRE meninjau kertas kerja EIR/CKPN dari konsultan","Dokumentasi peninjauan via email/approval, diketahui Kepala Divisi"],tanggapan_status:"proses",tanggapan_tgl:"20 Mar 2026",tanggapan_entries:["Sedang diskusi dengan konsultan VBOX untuk integrasi data realisasi pembayaran — target desain solusi Q2 2026","Peninjauan kertas kerja EIR/CKPN sudah diimplementasikan sejak Februari 2026 — template email approval sudah digunakan","Kepala Divisi AMK sudah memberikan approval via email untuk perhitungan bulan Feb & Mar 2026"]},
{ref:"ITAC-MKR-01",cycle:"Mekaar",type:"Otomatis",cause:"Ketidaklengkapan Data",desc:"Sistem Mekaar Digi memvalidasi data NIK/Nama/TTL dengan Dukcapil — data nasabah siklus lanjutan dapat diganti tanpa validasi",target:"Q1 2026",gap:"GAP-MKR-04",pic:"Divisi ATI",remediasi:["Menonaktifkan fitur input data pribadi pada siklus lanjutan","Meningkatkan pengecekan Dukcapil atas alamat, tempat lahir, KK"],tanggapan_status:"terkendala",tanggapan_tgl:"30 Mar 2026",tanggapan_entries:["Fitur edit data pribadi nasabah siklus lanjutan sudah dinonaktifkan pada Mekaar Digi versi 4.2.1 per 25 Maret 2026","Integrasi pengecekan Dukcapil untuk data tambahan (alamat, KK) TERKENDALA oleh keterbatasan API Dukcapil — sedang koordinasi langsung dengan Dukcapil"]},
{ref:"ITAC-MKR-13",cycle:"Mekaar",type:"Otomatis",cause:"Ketiadaan Bukti",desc:"Sistem Mekaar Digi memverifikasi kesesuaian nama nasabah dengan pemilik rekening — pencairan tetap berjalan meski nama berbeda",target:"Q2 2026",gap:"GAP-MKR-02",pic:"Divisi ATI & BUM",remediasi:["Mengembangkan blocker otomatis apabila nama pemilik rekening berbeda dengan nama nasabah"],tanggapan_status:"proses",tanggapan_tgl:"1 Apr 2026",tanggapan_entries:["Spesifikasi teknis blocker sudah difinalisasi bersama Divisi BUM","Pengembangan dimulai April 2026, target go-live Mei 2026"]},
{ref:"ITAC-MKR-22",cycle:"Mekaar",type:"Otomatis",cause:"Ketiadaan Bukti",desc:"Sistem BR.Net menghitung simulasi pelunasan dini — biaya jasa tidak terbentuk untuk pencairan di luar jadwal PKM",target:"Q1 2026",gap:"GAP-MKR-07",pic:"Divisi Bisnis & ATI",remediasi:["Divisi Bisnis skenario penyesuaian pelunasan dini di luar PKM","Divisi ATI implementasi hasil penyesuaian"],tanggapan_status:"proses",tanggapan_tgl:"1 Apr 2026",tanggapan_entries:["Skenario perhitungan pelunasan dini untuk pencairan di luar PKM sudah difinalisasi oleh Divisi Bisnis per 20 Maret 2026","Implementasi konfigurasi BR.Net sedang dilakukan oleh Divisi ATI, target selesai minggu ke-2 April 2026"]},
{ref:"ULM-81-02",cycle:"ULaMM",type:"Manual",cause:"Ketiadaan Bukti",desc:"Kepala Divisi AMK reviu perhitungan bunga EIR bulanan — bukti peninjauan CKPN/EIR konsultan belum terdokumentasi",target:"Q2 2026",gap:"GAP-ULM-06",pic:"Divisi AMK & MRE",remediasi:["Meninjau kertas kerja perhitungan EIR/CKPN","Dokumentasi peninjauan via email/approval","Hasil peninjauan diketahui Kepala Divisi"],tanggapan_status:"selesai",tanggapan_tgl:"20 Mar 2026",tanggapan_entries:["Seluruh item remediasi sudah diimplementasikan sejak Februari 2026","Template email approval sudah dibuat dan digunakan secara rutin","Kepala Divisi AMK sudah memberikan approval via email untuk perhitungan bulan Feb & Mar 2026"]},
{ref:"ITAC-ULM-02",cycle:"ULaMM",type:"Otomatis",cause:"Ketidakakuratan Informasi",desc:"Sistem Marketline validasi input bunga/plafon/tenor — belum sesuai standar BKU",target:"Q3 2026",gap:"GAP-ULM-01",pic:"Divisi ATI",remediasi:["Penyesuaian sistem Manpro/MT200 agar validasi:","a. Suku bunga sesuai jenjang plafon dan tenor","b. Jumlah tenor sesuai jenjang plafon"],tanggapan_status:"proses",tanggapan_tgl:"20 Mar 2026",tanggapan_entries:["Mapping parameter validasi per produk sedang dilakukan bersama Divisi BMK","Parameter bunga dan tenor akan disesuaikan setelah mapping selesai"]},
{ref:"ITAC-ULM-09",cycle:"ULaMM",type:"Otomatis",cause:"Ketidakakuratan Informasi",desc:"Sistem Marketline menentukan jalur approval BWMP — belum dibatasi sesuai pemangku jabatan terkini (nasabah 3R)",target:"Q2 2026",gap:"GAP-ULM-02",pic:"Divisi Kepatuhan & ATI",remediasi:["Implementasi MTDIGII untuk proses 3R","Otomatisasi update BWMP di MTDIGII"],tanggapan_status:"proses",tanggapan_tgl:"5 Apr 2026",tanggapan_entries:["MTDIGII untuk proses 3R (Reschedule, Recondition, Restructure) sedang dalam tahap UAT","Target go-live MTDIGII Mei 2026","Otomatisasi update BWMP menunggu go-live MTDIGII terlebih dahulu, target Q4 2026"]},
{ref:"FRP-11-01",cycle:"Fin. Reporting",type:"Manual",cause:"Ketiadaan Bukti",desc:"Wakil Kepala Divisi review persetujuan permohonan Kode Akun baru — tidak ada dokumentasi",target:"Mei 2026",gap:"GAP-FRP-15",pic:"Divisi AMK",remediasi:["Standarisasi proses pengajuan dan persetujuan COA baru agar terdokumentasi melalui email"],tanggapan_status:"selesai",tanggapan_tgl:"15 Mar 2026",tanggapan_entries:["SOP pengajuan dan persetujuan COA baru sudah diterbitkan","Dokumentasi via email sudah berjalan sejak Maret 2026"]},
{ref:"FRP-12-01",cycle:"Fin. Reporting",type:"Manual",cause:"Ketiadaan Bukti",desc:"Divisi AMK review mapping COA baru di core-system dengan COA Oracle — tidak ada dokumentasi",target:"Mei 2026",gap:"GAP-FRP-15",pic:"Divisi AMK",remediasi:["Standarisasi proses pengajuan dan persetujuan mapping COA agar terdokumentasi melalui email"],tanggapan_status:"selesai",tanggapan_tgl:"15 Mar 2026",tanggapan_entries:["Proses mapping COA sudah didokumentasikan via email sejak Maret 2026"]},
{ref:"FRP-41-02",cycle:"Fin. Reporting",type:"Manual",cause:"Ketiadaan Bukti",desc:"Kabag Akuntansi Operasional monitoring jurnal menggantung sebelum closing — tidak terdokumentasi",target:"Mei 2026",gap:"GAP-FRP-15",pic:"Divisi AMK",remediasi:["Standarisasi monitoring jurnal agar terdokumentasi melalui closing checklist","Mekanisme pelaksanaan diatur dalam kebijakan AMK"],tanggapan_status:"proses",tanggapan_tgl:"1 Apr 2026",tanggapan_entries:["Draft closing checklist sudah disusun, sedang dalam proses review oleh Wakadiv AMK","Draft kebijakan AMK terkait mekanisme pelaksanaan sedang disusun, target finalisasi Mei 2026"]},
{ref:"FRP-42-01",cycle:"Fin. Reporting",type:"Manual",cause:"Ketiadaan Bukti",desc:"Wakadiv AMK/Kabag persetujuan pembukaan kembali periode akuntansi Oracle — tidak terdokumentasi",target:"Mei 2026",gap:"GAP-FRP-15",pic:"Divisi AMK",remediasi:["Standarisasi proses pembukaan/penutupan periode agar terdokumentasi melalui email"],tanggapan_status:"selesai",tanggapan_tgl:"15 Mar 2026",tanggapan_entries:["Proses pengajuan pembukaan/penutupan periode sudah didokumentasikan via email sejak Maret 2026"]},
{ref:"FRP-43-01",cycle:"Fin. Reporting",type:"Manual",cause:"Ketiadaan Bukti",desc:"Kabag Pelaporan review kesesuaian TB Oracle post AJE vs TB Report BRI — tidak terdokumentasi",target:"Mei 2026",gap:"GAP-FRP-15",pic:"Divisi AMK",remediasi:["Standarisasi pengecekan TB agar terdokumentasi melalui closing checklist"],tanggapan_status:"proses",tanggapan_tgl:"1 Apr 2026",tanggapan_entries:["Akan dimasukkan ke dalam closing checklist yang sedang disusun"]},
{ref:"FRP-51-01",cycle:"Fin. Reporting",type:"Manual",cause:"Ketiadaan Bukti",desc:"Kabag, Wakadiv, Kadiv AMK review kertas kerja AJE — tidak terdokumentasi",target:"Mei 2026",gap:"GAP-FRP-15",pic:"Divisi AMK",remediasi:["Standarisasi review AJE agar terdokumentasi melalui closing checklist","Mekanisme diatur dalam kebijakan AMK"],tanggapan_status:"proses",tanggapan_tgl:"1 Apr 2026",tanggapan_entries:["Akan dimasukkan ke closing checklist yang sedang disusun","Draft kebijakan AMK sedang dalam proses penyusunan"]},
{ref:"FRP-52-01",cycle:"Fin. Reporting",type:"Manual",cause:"Ketiadaan Bukti",desc:"Kabag, Wakadiv, Kadiv AMK review eliminasi konsolidasi — tidak terdokumentasi",target:"Mei 2026",gap:"GAP-FRP-15",pic:"Divisi AMK",remediasi:["Standarisasi review eliminasi konsolidasi agar terdokumentasi melalui closing checklist","Mekanisme diatur dalam kebijakan AMK"],tanggapan_status:"proses",tanggapan_tgl:"1 Apr 2026",tanggapan_entries:["Akan dimasukkan ke closing checklist yang sedang disusun","Draft kebijakan sedang dalam proses penyusunan"]},
{ref:"ITAC-FRP-01",cycle:"Fin. Reporting",type:"Otomatis",cause:"Ketidaklengkapan Data",desc:"Akses modifikasi COA pada Oracle dibatasi untuk Admin Oracle AMK — 14 superuser aktif, 2 sudah pindah divisi",target:"Q1 2026",gap:"GAP-FRP-01",pic:"Divisi AMK & ATI",remediasi:["Profiling kembali Oracle Superuser agar modifikasi COA hanya 1 user","Pemisahan akses role AMK dan IT"],tanggapan_status:"proses",tanggapan_tgl:"1 Apr 2026",tanggapan_entries:["Profiling user Oracle telah dilakukan — akses modifikasi COA sudah dibatasi ke 1 user (Admin Oracle AMK) per 20 Maret 2026","Akses review sudah dibatasi ke Wakadiv AMK","Koordinasi dengan Divisi ATI untuk pemisahan role PNM KTRPST Superuser sedang berjalan"]},
{ref:"ITAC-FRP-02",cycle:"Fin. Reporting",type:"Otomatis",cause:"Ketidaklengkapan Data",desc:"Akses input dan posting jurnal tersegregasi pada Oracle — COA nihil/inaktif belum dimonitor periodik",target:"Q2 2026",gap:"GAP-FRP-07",pic:"Divisi AMK & TRP",remediasi:["Evaluasi periodik daftar rekening bank Oracle","Evaluasi tahunan COA Oracle","Konfirmasi PIC untuk COA Inaktif/Nihil","Membuat SOP monitoring COA"],tanggapan_status:"proses",tanggapan_tgl:"1 Apr 2026",tanggapan_entries:["Evaluasi pertama daftar rekening bank sudah dilakukan bulan Maret 2026 bersama Divisi TRP","Evaluasi tahunan COA Oracle dijadwalkan Q2 2026","Draft SOP monitoring COA sedang disusun oleh Divisi AMK"]},
{ref:"ITAC-FRP-06",cycle:"Fin. Reporting",type:"Otomatis",cause:"Ketidaklengkapan Data",desc:"Sistem BR.Net membatasi akses reversal jurnal — Divisi OPR posting tanpa batasan dan tanpa approval atasan",target:"Q1 2026",gap:"GAP-FRP-08",pic:"Divisi OPR & ATI",remediasi:["Update UAM BR.Net sesuai kondisi terkini","Limitasi per divisi: OPR=BAJTK+Kas; TRP=pencairan; CRO=klaim","Otomasi daily closing oleh STR untuk RAK otomatis"],tanggapan_status:"terkendala",tanggapan_tgl:"5 Apr 2026",tanggapan_entries:["UAM BR.Net sudah diperbarui dan didistribusikan ke seluruh pihak terkait per 28 Maret 2026","Penerapan limitasi per divisi TERKENDALA — memerlukan pengembangan fitur baru di BR.Net, sedang koordinasi dengan vendor","Inisiatif otomasi daily closing sedang dalam tahap desain oleh Divisi STR, target implementasi Q4 2026"]},
{ref:"ITAC-FRP-14",cycle:"Fin. Reporting",type:"Otomatis",cause:"Ketidaklengkapan Data",desc:"Akses open/close period pada Oracle dibatasi untuk Kabag — akses superuser belum terbatas",target:"Q1 2026",gap:"GAP-FRP-01",pic:"Divisi AMK & ATI",remediasi:["Membatasi akses open/close periode ke 2 user (Wakadiv AMK + Kabag)","Menghapus akses user yang sudah pindah divisi"],tanggapan_status:"selesai",tanggapan_tgl:"20 Mar 2026",tanggapan_entries:["Open/close periode akuntansi sudah dibatasi ke 2 user account per 20 Maret 2026","2 user yang sudah pindah divisi telah dinonaktifkan"]},
];

const DD={"2025":{label:"Tahun 2025",subtitle:"Piloting CSA | Des 2025 – Mar 2026",totalControls:270,totalTested:150,keyNotTested:5,totalEffective:131,totalIneffective:19,manualTotal:{efektif:87,inefektif:10},otomatisTotal:{efektif:44,inefektif:9},cycles:[{name:"Mekaar",total:92,tested:51,effective:46,ineffective:5,manualE:28,manualI:2,otomatisE:18,otomatisI:3},{name:"ULaMM",total:131,tested:72,effective:69,ineffective:3,manualE:50,manualI:1,otomatisE:19,otomatisI:2},{name:"Fin. Reporting",total:47,tested:27,effective:16,ineffective:11,manualE:9,manualI:7,otomatisE:7,otomatisI:4}],rootCauses:[{name:"Ketiadaan Bukti",count:8,color:"#ef4444"},{name:"Ketidaklengkapan Data",count:6,color:"#f59e0b"},{name:"Ketidakakuratan Informasi",count:5,color:"#8b5cf6"}],findings:FINDINGS,scope:{periode:"Des 2025 – Mar 2026",lokasi:"Kantor Pusat, Cabang, Unit",siklus:"Kredit Mekaar, Kredit ULaMM, Financial Reporting",divisi:"AMK, CRO, MRE, OPR, BMK, ATI, TRP, Cabang, Unit Mekaar, Unit ULaMM"}}};

const KPI=({label,value,sub,color})=>(<div style={{background:C.card,borderRadius:16,padding:"24px 20px",border:`1px solid ${C.border}`,flex:"1 1 150px",minWidth:140}}><div style={{fontSize:13,color:C.muted,marginBottom:6}}>{label}</div><div style={{fontSize:36,fontWeight:800,color:color||C.text,lineHeight:1.1}}>{value}</div>{sub&&<div style={{fontSize:12,color:C.muted,marginTop:4}}>{sub}</div>}</div>);
const CycleBar=({data})=>{const p=Math.round((data.effective/Math.max(data.tested,1))*100);return(<div style={{background:C.card,borderRadius:16,padding:20,border:`1px solid ${C.border}`}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><span style={{fontWeight:700,fontSize:15}}>{data.name}</span><span style={{fontSize:13,color:C.muted}}>{data.tested} diuji / {data.total} total</span></div><div style={{display:"flex",height:28,borderRadius:8,overflow:"hidden",background:"#1e293b"}}><div style={{width:`${p}%`,background:C.effective,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,color:"#fff"}}>{data.effective}</div><div style={{width:`${100-p}%`,background:C.ineffective,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,color:"#fff"}}>{data.ineffective}</div></div><div style={{display:"flex",gap:16,marginTop:10,fontSize:12,color:C.muted}}><span>Manual: {data.manualE}E / {data.manualI}I</span><span>Otomatis: {data.otomatisE}E / {data.otomatisI}I</span></div></div>);};
function PBar({status}){const pct=Math.min(100,Math.max(0,((180-Math.max(0,status.diff))/180)*100)),ov=status.diff<0;return(<div style={{width:"100%",marginTop:8}}><div style={{height:6,borderRadius:3,background:"#1e293b",overflow:"hidden"}}><div style={{height:"100%",borderRadius:3,width:`${ov?100:pct}%`,background:ov?`repeating-linear-gradient(90deg,${status.color},${status.color} 4px,${status.color}88 4px,${status.color}88 8px)`:`linear-gradient(90deg,${status.color}88,${status.color})`,transition:"width 0.8s"}}/></div><div style={{display:"flex",justifyContent:"space-between",marginTop:4,fontSize:11}}><span style={{color:status.color,fontWeight:600}}>{status.msg}</span><span style={{color:C.muted}}>Deadline: {fmtD(status.deadline)}</span></div></div>);}
const SBadge=({status})=>(<span style={{display:"inline-flex",alignItems:"center",gap:5,padding:"3px 10px",borderRadius:6,fontSize:11,fontWeight:700,background:status.bg,color:status.color,border:`1px solid ${status.color}44`,animation:status.p<=1?"pulse 2s infinite":"none"}}><span style={{width:6,height:6,borderRadius:"50%",background:status.color}}/>{status.label}</span>);
const TBadge=({s})=>{const t=TSTS[s]||TSTS.belum;return(<span style={{display:"inline-flex",alignItems:"center",gap:4,padding:"3px 10px",borderRadius:6,fontSize:11,fontWeight:700,background:t.bg,color:t.color,border:`1px solid ${t.color}33`}}>{t.icon} {t.label}</span>);};

export default function App(){
  const[periods,setPeriods]=useState(DD);
  const[period,setPeriod]=useState("2025");
  const[tab,setTab]=useState("overview");
  const[fCy,setFCy]=useState("All");
  const[fSt,setFSt]=useState("All");
  const[sortBy,setSortBy]=useState("urgency");
  const[expCard,setExpCard]=useState(null);
  const[ddO,setDdO]=useState(false);
  const[uMsg,setUMsg]=useState(null);
  const fRef=useRef(null);

  const handleUpload=useCallback((e)=>{const f=e.target.files?.[0];if(!f)return;const r=new FileReader();r.onload=(ev)=>{try{const res=parseXlsx(ev.target.result);setPeriods(p=>({...p,[res.id]:res.data}));setPeriod(res.id);setTab("overview");setUMsg({t:"s",m:`"${res.data.label}" berhasil di-upload!`});setTimeout(()=>setUMsg(null),4000)}catch(err){setUMsg({t:"e",m:`Gagal: ${err.message}`});setTimeout(()=>setUMsg(null),5000)}};r.readAsArrayBuffer(f);e.target.value=""},[]);

  const pd=periods[period]||periods[Object.keys(periods)[0]];
  const fws=useMemo(()=>(pd?.findings||[]).map(f=>({...f,status:getSts(f.target)})),[pd]);
  const sm=useMemo(()=>{const s={overdue:0,critical:0,warning:0,ontrack:0};fws.forEach(f=>{if(f.status.p===0)s.overdue++;else if(f.status.p===1)s.critical++;else if(f.status.p===2)s.warning++;else s.ontrack++});return s},[fws]);
  const tgStats=useMemo(()=>{let selesai=0,proses=0,belum=0,terkendala=0;fws.forEach(f=>{const s=f.tanggapan_status||"belum";if(s==="selesai")selesai++;else if(s==="proses")proses++;else if(s==="terkendala")terkendala++;else belum++});const total=fws.length;return{total,selesai,proses,belum,terkendala,pct:total?Math.round((selesai/total)*100):0}},[fws]);

  const eRate=pd?.totalTested>0?Math.round((pd.totalEffective/pd.totalTested)*100):0;
  const uCy=[...new Set((pd?.findings||[]).map(f=>f.cycle))];
  let filtered=fCy==="All"?fws:fws.filter(d=>d.cycle===fCy);
  if(fSt!=="All")filtered=filtered.filter(d=>d.status.label===fSt);
  if(sortBy==="urgency")filtered=[...filtered].sort((a,b)=>a.status.p-b.status.p||a.status.diff-b.status.diff);
  else if(sortBy==="deadline")filtered=[...filtered].sort((a,b)=>a.status.diff-b.status.diff);
  else filtered=[...filtered].sort((a,b)=>a.cycle.localeCompare(b.cycle));

  const tabs=[{id:"overview",label:"Overview"},{id:"findings",label:"Tracking"},{id:"detail",label:"Detail Remediasi"},{id:"tanggapan",label:"Tanggapan Lini 1"}];

  return(<div style={{background:C.bg,minHeight:"100vh",color:C.text,fontFamily:"'DM Sans',sans-serif",padding:"24px 20px"}} onClick={()=>ddO&&setDdO(false)}>
    <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700;800&display=swap" rel="stylesheet"/>
    <style>{`@keyframes pulse{0%,100%{opacity:1}50%{opacity:.6}}`}</style>
    <input type="file" ref={fRef} accept=".xlsx,.xls" onChange={handleUpload} style={{display:"none"}}/>
    {uMsg&&<div style={{position:"fixed",top:16,right:16,zIndex:100,padding:"12px 20px",borderRadius:12,background:uMsg.t==="s"?"#22c55e22":"#ef444422",border:`1px solid ${uMsg.t==="s"?"#22c55e":"#ef4444"}66`,color:uMsg.t==="s"?"#22c55e":"#ef4444",fontSize:13,fontWeight:600,boxShadow:"0 8px 32px rgba(0,0,0,0.3)"}}>{uMsg.t==="s"?"✅":"❌"} {uMsg.m}</div>}

    {/* Header */}
    <div style={{marginBottom:28}}><div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
      <div style={{display:"flex",alignItems:"center",gap:12}}><div style={{width:8,height:32,borderRadius:4,background:"linear-gradient(180deg,#38bdf8,#6366f1)"}}/><div><h1 style={{fontSize:22,fontWeight:800,margin:0}}>CSA ICoFR Dashboard</h1><div style={{fontSize:13,color:C.muted}}>PT Permodalan Nasional Madani — {pd?.subtitle}</div></div></div>
      <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
        <button onClick={()=>fRef.current?.click()} style={{display:"flex",alignItems:"center",gap:6,padding:"10px 16px",background:"linear-gradient(135deg,#22c55e33,#22c55e11)",border:"1px solid #22c55e44",borderRadius:12,color:"#22c55e",cursor:"pointer",fontSize:13,fontWeight:600}}>↑ Upload Excel</button>
        <div style={{position:"relative"}}><button onClick={e=>{e.stopPropagation();setDdO(!ddO)}} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 18px",background:C.card,border:`1px solid ${C.border}`,borderRadius:12,color:C.text,cursor:"pointer",fontSize:14,fontWeight:600}}><span style={{width:10,height:10,borderRadius:"50%",background:C.effective}}/>{pd?.label}<svg width="12" height="12" viewBox="0 0 12 12" fill="none" style={{transform:ddO?"rotate(180deg)":"none",transition:"transform 0.2s"}}><path d="M3 4.5L6 7.5L9 4.5" stroke={C.muted} strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/></svg></button>
        {ddO&&<div onClick={e=>e.stopPropagation()} style={{position:"absolute",top:"calc(100% + 6px)",right:0,zIndex:50,background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:6,minWidth:260,boxShadow:"0 16px 48px rgba(0,0,0,0.5)"}}>{Object.entries(periods).map(([k,v])=>(<button key={k} onClick={()=>{setPeriod(k);setDdO(false)}} style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"10px 14px",border:"none",borderRadius:8,cursor:"pointer",background:period===k?`${C.accent}18`:"transparent",color:period===k?C.accent:C.text,fontSize:13,fontWeight:period===k?700:500,textAlign:"left"}}><span style={{width:8,height:8,borderRadius:"50%",background:C.effective,flexShrink:0}}/><div style={{flex:1}}><div>{v.label}</div><div style={{fontSize:11,color:C.muted,fontWeight:400}}>{v.subtitle}</div></div>{period===k&&<span>✓</span>}</button>))}</div>}</div>
      </div>
    </div></div>

    {/* Tabs */}
    <div style={{display:"flex",gap:4,marginBottom:24,background:C.card,borderRadius:12,padding:4,border:`1px solid ${C.border}`,flexWrap:"wrap",width:"fit-content"}}>
      {tabs.map(t=>(<button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"8px 16px",borderRadius:10,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:tab===t.id?"linear-gradient(135deg,#38bdf8,#6366f1)":"transparent",color:tab===t.id?"#fff":C.muted,position:"relative",whiteSpace:"nowrap"}}>
        {t.label}
        {t.id==="findings"&&sm.overdue>0&&<span style={{position:"absolute",top:-4,right:-4,width:18,height:18,borderRadius:"50%",background:"#ef4444",color:"#fff",fontSize:10,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center",animation:"pulse 1.5s infinite"}}>{sm.overdue}</span>}
        {t.id==="tanggapan"&&tgStats.terkendala>0&&<span style={{position:"absolute",top:-4,right:-4,width:18,height:18,borderRadius:"50%",background:"#f97316",color:"#fff",fontSize:10,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center"}}>{tgStats.terkendala}</span>}
      </button>))}
    </div>

    {/* OVERVIEW */}
    {tab==="overview"&&(<>
      {sm.overdue>0&&<div style={{background:"#ef444418",border:"1px solid #ef444444",borderRadius:12,padding:"14px 20px",marginBottom:20,display:"flex",alignItems:"center",gap:12,cursor:"pointer"}} onClick={()=>{setTab("findings");setFSt("OVERDUE")}}><span style={{fontSize:22,animation:"pulse 1.5s infinite"}}>⚠️</span><div><div style={{fontSize:14,fontWeight:700,color:"#ef4444"}}>{sm.overdue} remediasi telah melewati deadline</div></div></div>}
      <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:24}}><KPI label="Total Pengendalian" value={pd.totalControls} sub="TLC teridentifikasi"/><KPI label="Key Control Diuji" value={pd.totalTested} sub={`dari ${pd.totalTested+pd.keyNotTested} key controls`}/><KPI label="Efektif" value={pd.totalEffective} sub={`${eRate}% efektivitas`} color={C.effective}/><KPI label="Inefektif" value={pd.totalIneffective} sub="perlu remediasi" color={C.ineffective}/></div>
      <div style={{display:"flex",gap:16,marginBottom:20,flexWrap:"wrap"}}>
        <div style={{background:C.card,borderRadius:16,padding:20,border:`1px solid ${C.border}`,flex:"1 1 280px",minWidth:280}}><div style={{fontSize:14,fontWeight:700,marginBottom:12}}>Hasil CSA</div><ResponsiveContainer width="100%" height={200}><PieChart><Pie data={[{name:"Efektif",value:pd.totalEffective},{name:"Inefektif",value:pd.totalIneffective}]} cx="50%" cy="50%" innerRadius={55} outerRadius={85} paddingAngle={4} dataKey="value" stroke="none"><Cell fill={C.effective}/><Cell fill={C.ineffective}/></Pie><Tooltip contentStyle={{background:C.card,border:`1px solid ${C.border}`,borderRadius:8,color:C.text,fontSize:13}}/><Legend wrapperStyle={{fontSize:12}}/></PieChart></ResponsiveContainer><div style={{textAlign:"center",marginTop:-8,fontSize:28,fontWeight:800,color:C.effective}}>{eRate}%</div></div>
        <div style={{background:C.card,borderRadius:16,padding:20,border:`1px solid ${C.border}`,flex:"2 1 400px",minWidth:320}}><div style={{fontSize:14,fontWeight:700,marginBottom:12}}>Hasil per Siklus</div><ResponsiveContainer width="100%" height={220}><BarChart data={pd.cycles.map(c=>({name:c.name,Efektif:c.effective,Inefektif:c.ineffective}))} barGap={4}><XAxis dataKey="name" tick={{fill:C.muted,fontSize:12}} axisLine={false} tickLine={false}/><YAxis tick={{fill:C.muted,fontSize:12}} axisLine={false} tickLine={false}/><Tooltip contentStyle={{background:C.card,border:`1px solid ${C.border}`,borderRadius:8,color:C.text,fontSize:13}}/><Legend wrapperStyle={{fontSize:12}}/><Bar dataKey="Efektif" fill={C.effective} radius={[6,6,0,0]}/><Bar dataKey="Inefektif" fill={C.ineffective} radius={[6,6,0,0]}/></BarChart></ResponsiveContainer></div>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:12,marginBottom:20}}>{pd.cycles.map(c=><CycleBar key={c.name} data={c}/>)}</div>
      <div style={{display:"flex",gap:16,flexWrap:"wrap"}}>
        <div style={{background:C.card,borderRadius:16,padding:20,border:`1px solid ${C.border}`,flex:"1 1 320px"}}><div style={{fontSize:14,fontWeight:700,marginBottom:12}}>Manual vs Otomatis</div><ResponsiveContainer width="100%" height={200}><BarChart data={[{name:"Manual",Efektif:pd.manualTotal.efektif,Inefektif:pd.manualTotal.inefektif},{name:"Otomatis",Efektif:pd.otomatisTotal.efektif,Inefektif:pd.otomatisTotal.inefektif}]} layout="vertical" barGap={4}><XAxis type="number" tick={{fill:C.muted,fontSize:12}} axisLine={false} tickLine={false}/><YAxis dataKey="name" type="category" tick={{fill:C.muted,fontSize:12}} axisLine={false} tickLine={false} width={70}/><Tooltip contentStyle={{background:C.card,border:`1px solid ${C.border}`,borderRadius:8,color:C.text,fontSize:13}}/><Legend wrapperStyle={{fontSize:12}}/><Bar dataKey="Efektif" fill={C.effective} radius={[0,6,6,0]} barSize={24}/><Bar dataKey="Inefektif" fill={C.ineffective} radius={[0,6,6,0]} barSize={24}/></BarChart></ResponsiveContainer></div>
        <div style={{background:C.card,borderRadius:16,padding:20,border:`1px solid ${C.border}`,flex:"1 1 320px"}}><div style={{fontSize:14,fontWeight:700,marginBottom:16}}>Penyebab Inefektivitas</div>{pd.rootCauses.map(rc=>(<div key={rc.name} style={{marginBottom:16}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}><span style={{fontSize:13}}>{rc.name}</span><span style={{fontSize:13,fontWeight:700,color:rc.color}}>{rc.count}</span></div><div style={{height:8,borderRadius:4,background:"#1e293b"}}><div style={{height:"100%",borderRadius:4,width:`${(rc.count/Math.max(pd.totalIneffective,1))*100}%`,background:rc.color}}/></div></div>))}</div>
      </div>
    </>)}

    {/* TRACKING */}
    {tab==="findings"&&(<>
      <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:20}}>{[{k:"overdue",...STS.overdue,count:sm.overdue,sub:"Lewat deadline"},{k:"critical",...STS.critical,count:sm.critical,sub:"< 30 hari"},{k:"warning",...STS.warning,count:sm.warning,sub:"< 90 hari"},{k:"ontrack",...STS.ontrack,count:sm.ontrack,sub:"> 90 hari"}].map(s=>(<div key={s.k} onClick={()=>setFSt(fSt===s.label?"All":s.label)} style={{background:fSt===s.label?s.bg:C.card,borderRadius:12,padding:"14px 20px",border:`1px solid ${fSt===s.label?s.color+"66":C.border}`,cursor:"pointer",flex:"1 1 120px",minWidth:120,textAlign:"center"}}><div style={{fontSize:28,fontWeight:800,color:s.color}}>{s.count}</div><div style={{fontSize:11,fontWeight:700,color:s.color,marginTop:4}}>{s.label}</div><div style={{fontSize:11,color:C.muted,marginTop:2}}>{s.sub}</div></div>))}</div>
      <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}><span style={{fontSize:12,color:C.muted,fontWeight:600}}>Siklus:</span>{["All",...uCy].map(f=>(<button key={f} onClick={()=>setFCy(f)} style={{padding:"5px 14px",borderRadius:8,border:`1px solid ${fCy===f?C.accent:C.border}`,background:fCy===f?`${C.accent}22`:"transparent",color:fCy===f?C.accent:C.muted,cursor:"pointer",fontSize:12,fontWeight:600}}>{f==="All"?"Semua":f}</button>))}<div style={{flex:1}}/><span style={{fontSize:12,color:C.muted,fontWeight:600}}>Sort:</span>{[{k:"urgency",l:"Urgensi"},{k:"deadline",l:"Deadline"},{k:"cycle",l:"Siklus"}].map(s=>(<button key={s.k} onClick={()=>setSortBy(s.k)} style={{padding:"5px 12px",borderRadius:8,border:`1px solid ${sortBy===s.k?C.accent:C.border}`,background:sortBy===s.k?`${C.accent}22`:"transparent",color:sortBy===s.k?C.accent:C.muted,cursor:"pointer",fontSize:12,fontWeight:600}}>{s.l}</button>))}</div>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>{filtered.map((item,i)=>(<div key={i} style={{background:C.card,borderRadius:14,padding:"18px 20px",border:`1px solid ${C.border}`,borderLeft:`4px solid ${item.status.color}`}}><div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:8}}><div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}><span style={{fontSize:14,fontWeight:700,color:C.accent,fontFamily:"monospace"}}>{item.ref}</span><span style={{fontSize:11,padding:"2px 8px",borderRadius:4,background:`${cc(item.cycle)}33`,color:cc(item.cycle),fontWeight:600}}>{item.cycle}</span><SBadge status={item.status}/></div><span style={{fontSize:11,padding:"2px 6px",borderRadius:4,background:"#1e293b",color:C.muted,fontFamily:"monospace"}}>{item.gap}</span></div><div style={{fontSize:13,color:C.text,lineHeight:1.5,marginBottom:8}}>{item.desc}</div><div style={{display:"flex",gap:16,flexWrap:"wrap",fontSize:12,marginBottom:4}}><span><span style={{color:C.muted}}>PIC: </span><span style={{fontWeight:600,color:C.accent}}>{item.pic}</span></span><span><span style={{color:C.muted}}>Target: </span><span style={{fontWeight:600}}>{item.target}</span></span><TBadge s={item.tanggapan_status}/></div><PBar status={item.status}/></div>))}</div>
    </>)}

    {/* DETAIL REMEDIASI */}
    {tab==="detail"&&(<>
      <div style={{fontSize:14,color:C.muted,marginBottom:20}}>Detail rencana tindak lanjut yang disepakati Manajemen — {pd.label}.</div>
      {uCy.map(cy=>{const items=fws.filter(f=>f.cycle===cy);if(!items.length)return null;return(<div key={cy} style={{marginBottom:28}}><div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}><div style={{width:4,height:24,borderRadius:2,background:cc(cy)}}/><h2 style={{fontSize:17,fontWeight:800,margin:0}}>Siklus — {cy}</h2></div>
      <div style={{display:"flex",flexDirection:"column",gap:12}}>{items.map((item,i)=>{const isE=expCard===`d-${cy}-${i}`;return(<div key={i} style={{background:C.card,borderRadius:14,border:`1px solid ${C.border}`,overflow:"hidden"}}><div onClick={()=>setExpCard(isE?null:`d-${cy}-${i}`)} style={{padding:"18px 20px",cursor:"pointer",borderLeft:`4px solid ${cc(cy)}`}}><div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8}}><div style={{flex:1}}><div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",marginBottom:6}}><span style={{fontSize:14,fontWeight:700,color:C.accent,fontFamily:"monospace"}}>{item.ref}</span><SBadge status={item.status}/><span style={{fontSize:11,padding:"2px 6px",borderRadius:4,background:"#1e293b",color:C.muted,fontFamily:"monospace"}}>{item.gap}</span></div><div style={{fontSize:13,color:C.text,lineHeight:1.5}}>{item.desc}</div></div><svg width="20" height="20" viewBox="0 0 20 20" fill="none" style={{transform:isE?"rotate(180deg)":"none",transition:"transform 0.3s",flexShrink:0}}><path d="M5 8L10 13L15 8" stroke={C.muted} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg></div><PBar status={item.status}/></div>
      {isE&&<div style={{padding:"16px 20px 20px 24px",borderTop:`1px solid ${C.border}`,background:"#0d1321"}}><div style={{fontSize:14,fontWeight:700,color:C.accent,marginBottom:14}}>⭐ Rencana Tindak Lanjut</div>{(item.remediasi||[]).map((r,j)=>{const isSub=r.startsWith("a.")||r.startsWith("b.")||r.startsWith("c.");return(<div key={j} style={{display:"flex",gap:10,marginLeft:isSub?28:0,marginBottom:8}}>{!isSub&&<div style={{width:24,height:24,borderRadius:8,background:`${C.accent}22`,border:`1px solid ${C.accent}44`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,color:C.accent,flexShrink:0}}>{j+1}</div>}{isSub&&<div style={{width:6,height:6,borderRadius:"50%",background:C.muted,flexShrink:0,marginTop:7}}/>}<div style={{fontSize:13,color:C.text,lineHeight:1.6}}>{isSub?r.slice(3):r}</div></div>)})}</div>}</div>)})}</div></div>)})}
    </>)}

    {/* TANGGAPAN LINI 1 */}
    {tab==="tanggapan"&&(<>
      {/* Summary */}
      <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:20}}>
        <div style={{background:C.card,borderRadius:14,padding:"18px 24px",border:`1px solid ${C.border}`,flex:"1 1 200px"}}>
          <div style={{fontSize:13,color:C.muted,marginBottom:4}}>Overall Progres</div>
          <div style={{fontSize:40,fontWeight:800,color:tgStats.pct===100?C.effective:C.accent}}>{tgStats.pct}%</div>
          <div style={{height:8,borderRadius:4,background:"#1e293b",marginTop:8}}><div style={{height:"100%",borderRadius:4,background:`linear-gradient(90deg,${C.accent},${C.effective})`,width:`${tgStats.pct}%`}}/></div>
          <div style={{fontSize:11,color:C.muted,marginTop:6}}>{tgStats.selesai} dari {tgStats.total} temuan selesai</div>
        </div>
        {[{k:"selesai",...TSTS.selesai,c:tgStats.selesai},{k:"proses",...TSTS.proses,c:tgStats.proses},{k:"terkendala",...TSTS.terkendala,c:tgStats.terkendala},{k:"belum",...TSTS.belum,c:tgStats.belum}].map(s=>(<div key={s.k} style={{background:C.card,borderRadius:14,padding:"14px 20px",border:`1px solid ${C.border}`,flex:"1 1 100px",minWidth:100,textAlign:"center"}}><div style={{fontSize:24,marginBottom:2}}>{s.icon}</div><div style={{fontSize:28,fontWeight:800,color:s.color}}>{s.c}</div><div style={{fontSize:11,color:C.muted,marginTop:2}}>{s.label}</div></div>))}
      </div>

      <div style={{fontSize:14,color:C.muted,marginBottom:20,lineHeight:1.6}}>Tanggapan dari Lini 1 (control owner) atas rencana tindak lanjut remediasi. Tanggapan bisa terus di-update tanpa batas.</div>

      {uCy.map(cy=>{const items=fws.filter(f=>f.cycle===cy);if(!items.length)return null;
      return(<div key={cy} style={{marginBottom:28}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}><div style={{width:4,height:24,borderRadius:2,background:cc(cy)}}/><h2 style={{fontSize:17,fontWeight:800,margin:0}}>Siklus — {cy}</h2></div>
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          {items.map((item,i)=>{
            const isE=expCard===`t-${cy}-${i}`;
            const ts=TSTS[item.tanggapan_status]||TSTS.belum;
            const hasIssue=item.tanggapan_status==="terkendala";
            const entries=item.tanggapan_entries||[];

            return(<div key={i} style={{background:C.card,borderRadius:14,border:`1px solid ${hasIssue?"#f9731644":C.border}`,overflow:"hidden"}}>
              <div onClick={()=>setExpCard(isE?null:`t-${cy}-${i}`)} style={{padding:"18px 20px",cursor:"pointer",borderLeft:`4px solid ${cc(cy)}`}}>
                <div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
                  <div style={{flex:1}}>
                    <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",marginBottom:6}}>
                      <span style={{fontSize:14,fontWeight:700,color:C.accent,fontFamily:"monospace"}}>{item.ref}</span>
                      <SBadge status={item.status}/>
                      <TBadge s={item.tanggapan_status}/>
                      {hasIssue&&<span style={{fontSize:11,padding:"2px 8px",borderRadius:4,background:TSTS.terkendala.bg,color:TSTS.terkendala.color,fontWeight:700,border:"1px solid #f9731633"}}>⚠️ Ada Kendala</span>}
                    </div>
                    <div style={{fontSize:13,color:C.text,lineHeight:1.5,marginBottom:6}}>{item.desc}</div>
                    <div style={{display:"flex",gap:16,flexWrap:"wrap",fontSize:12}}>
                      <span><span style={{color:C.muted}}>PIC: </span><span style={{fontWeight:600,color:C.accent}}>{item.pic}</span></span>
                      {item.tanggapan_tgl&&<span><span style={{color:C.muted}}>Update terakhir: </span><span style={{fontWeight:600}}>{item.tanggapan_tgl}</span></span>}
                      <span><span style={{color:C.muted}}>Tanggapan: </span><span style={{fontWeight:600}}>{entries.length} catatan</span></span>
                    </div>
                  </div>
                  <svg width="20" height="20" viewBox="0 0 20 20" fill="none" style={{transform:isE?"rotate(180deg)":"none",transition:"transform 0.3s",flexShrink:0,marginTop:4}}><path d="M5 8L10 13L15 8" stroke={C.muted} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>
                </div>
              </div>

              {isE&&<div style={{borderTop:`1px solid ${C.border}`,background:"#0d1321",padding:"20px 20px 20px 24px"}}>
                {/* Remediasi summary */}
                <div style={{marginBottom:20}}>
                  <div style={{fontSize:13,fontWeight:700,color:C.accent,marginBottom:10}}>📋 Rencana Remediasi</div>
                  {(item.remediasi||[]).map((r,j)=>{const isSub=r.startsWith("a.")||r.startsWith("b.")||r.startsWith("c.");return(<div key={j} style={{display:"flex",gap:8,marginLeft:isSub?24:0,marginBottom:4}}>{!isSub&&<span style={{color:C.accent,fontWeight:700,fontSize:12,flexShrink:0}}>{j+1}.</span>}{isSub&&<span style={{color:C.muted,fontSize:12}}>•</span>}<span style={{fontSize:13,color:C.muted,lineHeight:1.5}}>{isSub?r.slice(3):r}</span></div>)})}
                </div>

                {/* Tanggapan */}
                <div>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
                    <span style={{fontSize:16}}>{ts.icon}</span>
                    <span style={{fontSize:14,fontWeight:700,color:ts.color}}>Tanggapan Lini 1</span>
                    <TBadge s={item.tanggapan_status}/>
                    {item.tanggapan_tgl&&<span style={{fontSize:11,color:C.muted,marginLeft:8}}>Update: {item.tanggapan_tgl}</span>}
                  </div>

                  {entries.length>0?(
                    <div style={{display:"flex",flexDirection:"column",gap:8}}>
                      {entries.map((entry,j)=>(
                        <div key={j} style={{display:"flex",gap:10}}>
                          <div style={{width:2,borderRadius:1,background:ts.color,flexShrink:0,opacity:0.5}}/>
                          <div style={{fontSize:13,color:C.text,lineHeight:1.6}}>{entry}</div>
                        </div>
                      ))}
                    </div>
                  ):(
                    <div style={{fontSize:13,color:C.muted,fontStyle:"italic",padding:"12px 16px",background:"#1e293b44",borderRadius:8}}>Belum ada tanggapan dari Lini 1</div>
                  )}
                </div>

                {/* Alert for terkendala */}
                {hasIssue&&<div style={{marginTop:16,padding:"12px 16px",background:"#f9731615",borderRadius:10,border:"1px solid #f9731633",display:"flex",alignItems:"center",gap:10}}>
                  <span style={{fontSize:16}}>⚠️</span>
                  <div><div style={{fontSize:12,fontWeight:700,color:"#f97316"}}>Item ini terkendala — perlu eskalasi/follow-up dari Divisi MR</div><div style={{fontSize:11,color:C.muted}}>Koordinasi dengan {item.pic}</div></div>
                </div>}
              </div>}
            </div>);
          })}
        </div>
      </div>)})}
    </>)}
  </div>);
}
