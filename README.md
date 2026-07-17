[index.html](https://github.com/user-attachments/files/30134159/index.html)
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>FinTrack</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0f1117;--bg2:#171b26;--bg3:#1e2333;--bg4:#252b3d;
  --border:rgba(255,255,255,0.07);--border2:rgba(255,255,255,0.13);
  --text:#f0f2f8;--text2:#8b93a8;--text3:#5a6280;
  --green:#22d3a0;--green-bg:rgba(34,211,160,0.1);
  --red:#f56565;--red-bg:rgba(245,101,101,0.1);
  --blue:#60a5fa;--amber:#fbbf24;--purple:#a78bfa;
  --radius:12px;--radius-sm:8px;--sidebar:220px;
}
body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;font-size:14px;line-height:1.6;-webkit-tap-highlight-color:transparent}
.app{display:flex;min-height:100vh}
.sidebar{width:var(--sidebar);background:var(--bg2);border-right:1px solid var(--border);padding:24px 0;display:flex;flex-direction:column;position:fixed;top:0;left:0;height:100vh;z-index:100;transition:transform .25s}
.logo{padding:0 20px 24px;font-size:16px;font-weight:600;letter-spacing:-.3px;display:flex;align-items:center;gap:8px;border-bottom:1px solid var(--border);margin-bottom:16px;flex-shrink:0}
.logo-icon{width:30px;height:30px;background:linear-gradient(135deg,var(--green),#0ea5e9);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0}
.sidebar-nav{display:flex;flex-direction:column;flex:1;overflow-y:auto;overflow-x:hidden}
.nav-item{display:flex;align-items:center;gap:10px;padding:10px 20px;cursor:pointer;transition:background .15s;color:var(--text2);font-size:13.5px;font-weight:400;border-left:2px solid transparent;user-select:none;flex-shrink:0}
.nav-item:hover{background:var(--bg3);color:var(--text)}
.nav-item.active{background:var(--bg3);color:var(--green);border-left-color:var(--green)}
.nav-spacer{flex:1}
.nav-icon{font-size:16px;width:20px;text-align:center;flex-shrink:0}
.main{margin-left:var(--sidebar);flex:1;padding:28px 32px;max-width:1280px}
.page-header{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:24px;flex-wrap:wrap;gap:12px}
.page-title{font-size:21px;font-weight:600;letter-spacing:-.4px}
.page-sub{font-size:13px;color:var(--text2);margin-top:2px}
.hdr-actions{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.cards-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:24px}
.card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:18px 20px}
.card-label{font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.6px;margin-bottom:6px}
.card-value{font-size:22px;font-weight:600;font-family:'DM Mono',monospace;letter-spacing:-.5px}
.card-sub{font-size:12px;color:var(--text2);margin-top:3px}
.content-grid{display:grid;grid-template-columns:1fr 360px;gap:18px;align-items:start}
.chart-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:22px}
.chart-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:8px}
.chart-title{font-size:14px;font-weight:500}
.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--text2)}
.legend-item{display:flex;align-items:center;gap:5px}
.legend-dot{width:8px;height:8px;border-radius:2px;flex-shrink:0}
.transactions-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden}
.tx-header{display:flex;align-items:center;justify-content:space-between;padding:16px 18px 14px;border-bottom:1px solid var(--border);flex-wrap:wrap;gap:8px}
.tx-title{font-size:14px;font-weight:500}
.tx-list{overflow-y:auto;max-height:420px}
.tx-list::-webkit-scrollbar{width:4px}
.tx-list::-webkit-scrollbar-thumb{background:var(--border2);border-radius:4px}
.tx-item{display:flex;align-items:center;gap:11px;padding:11px 18px;border-bottom:1px solid var(--border);transition:background .1s}
.tx-item:hover{background:var(--bg3)}
.tx-item:last-child{border-bottom:none}
.tx-icon{width:34px;height:34px;border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0}
.tx-info{flex:1;min-width:0}
.tx-name{font-size:13px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tx-cat{font-size:11px;color:var(--text2)}
.tx-date{font-size:11px;color:var(--text3);white-space:nowrap;margin-right:4px}
.tx-amount{font-family:'DM Mono',monospace;font-size:13px;font-weight:500;white-space:nowrap}
.tx-actions{display:flex;gap:2px;opacity:0;transition:opacity .15s}
.tx-item:hover .tx-actions{opacity:1}
.tx-btn{background:none;border:none;color:var(--text3);cursor:pointer;font-size:14px;padding:3px 6px;border-radius:5px;transition:all .15s}
.tx-btn:hover{background:var(--bg4)}
.btn{padding:7px 14px;border-radius:var(--radius-sm);border:1px solid var(--border2);background:var(--bg3);color:var(--text);font-size:12px;font-weight:500;cursor:pointer;font-family:'DM Sans',sans-serif;transition:all .15s;display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
.btn:hover{background:var(--bg4)}
.btn-primary{background:var(--green);color:#0f1117;border-color:var(--green)}
.btn-primary:hover{background:#1ab88c}
.btn-danger{background:var(--red-bg);color:var(--red);border-color:rgba(245,101,101,.3)}
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:200;align-items:center;justify-content:center;padding:16px}
.modal-bg.open{display:flex}
.modal{background:var(--bg2);border:1px solid var(--border2);border-radius:var(--radius);padding:26px;width:460px;max-width:100%;max-height:90vh;overflow-y:auto}
.modal-title{font-size:15px;font-weight:600;margin-bottom:18px}
.form-row{margin-bottom:13px}
.form-label{font-size:11.5px;color:var(--text2);margin-bottom:5px;display:block}
.form-input,.form-select{width:100%;padding:9px 11px;background:var(--bg3);border:1px solid var(--border2);border-radius:var(--radius-sm);color:var(--text);font-family:'DM Sans',sans-serif;font-size:13px;outline:none;transition:border .15s}
.form-input:focus,.form-select:focus{border-color:var(--green)}
.form-select option{background:var(--bg3)}
.form-row-2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.radio-group{display:flex;gap:8px}
.radio-btn{flex:1;padding:8px 10px;border-radius:var(--radius-sm);border:1px solid var(--border2);cursor:pointer;text-align:center;font-size:13px;font-weight:500;transition:all .15s;color:var(--text2);background:var(--bg3)}
.radio-btn.sel-income{background:var(--green-bg);border-color:var(--green);color:var(--green)}
.radio-btn.sel-expense{background:var(--red-bg);border-color:var(--red);color:var(--red)}
.modal-actions{display:flex;gap:10px;justify-content:flex-end;margin-top:18px}
.period-select{padding:7px 11px;background:var(--bg3);border:1px solid var(--border);border-radius:var(--radius-sm);color:var(--text);font-size:13px;font-family:'DM Sans',sans-serif;cursor:pointer;outline:none}
/* Filter bar with selectable chips */
.filter-bar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px;align-items:center}
.filter-chip{padding:5px 12px;border-radius:20px;border:1px solid var(--border);font-size:12px;cursor:pointer;background:var(--bg2);color:var(--text2);transition:all .15s;user-select:none;display:flex;align-items:center;gap:5px}
.filter-chip:hover{border-color:var(--border2);color:var(--text)}
.filter-chip.active-income{background:var(--green-bg);color:var(--green);border-color:var(--green)}
.filter-chip.active-expense{background:var(--red-bg);color:var(--red);border-color:var(--red)}
.filter-chip.active{background:var(--bg4);color:var(--text);border-color:var(--border2)}
.filter-chip .chip-x{font-size:10px;opacity:.6;margin-left:2px}
.filter-chip .chip-x:hover{opacity:1}
.filter-section{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:16px 18px;margin-bottom:16px}
.filter-section-title{font-size:12px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px}
.chip-group{display:flex;gap:6px;flex-wrap:wrap}
/* Cat bars */
.cat-list{padding:6px 0}
.cat-row{display:flex;align-items:center;gap:11px;padding:9px 20px}
.cat-bar-wrap{flex:1;height:5px;background:var(--bg3);border-radius:3px;overflow:hidden}
.cat-bar{height:100%;border-radius:3px;transition:width .4s ease}
.cat-name-col{font-size:13px;width:130px;color:var(--text2);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cat-amt{font-family:'DM Mono',monospace;font-size:12px;color:var(--text);min-width:90px;text-align:right}
.cat-pct{font-size:11px;color:var(--text3);min-width:34px;text-align:right}
.cat-manager{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.cat-section-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:20px}
.cat-section-title{font-size:14px;font-weight:500;margin-bottom:14px}
.cat-item-row{display:flex;align-items:center;gap:9px;padding:7px 0;border-bottom:1px solid var(--border)}
.cat-item-row:last-child{border-bottom:none}
.cat-color-dot{width:11px;height:11px;border-radius:50%;flex-shrink:0}
.cat-item-name{flex:1;font-size:13px}
.cat-item-badge{font-size:10px;padding:2px 7px;border-radius:10px;background:var(--bg3);color:var(--text3)}
.cat-item-edit-btn{background:none;border:none;color:var(--text3);cursor:pointer;font-size:13px;opacity:.5;padding:2px 6px}
.cat-item-edit-btn:hover{color:var(--blue);opacity:1}
.cat-edit-inline{background:var(--bg3);border:1px solid var(--border2);border-radius:var(--radius-sm);padding:10px 12px;margin-top:6px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.cat-total-badge{font-size:11px;font-family:'DM Mono',monospace;padding:2px 8px;border-radius:20px;background:var(--bg3);color:var(--text2)}
.cat-item-del{background:none;border:none;color:var(--text3);cursor:pointer;font-size:13px;opacity:.5;padding:2px 6px}
.cat-item-del:hover{color:var(--red);opacity:1}
.new-cat-form{display:flex;gap:7px;margin-top:14px;flex-wrap:wrap}
.color-picker{width:38px;height:34px;border:1px solid var(--border2);border-radius:var(--radius-sm);cursor:pointer;background:none;padding:2px}
.drop-zone{border:2px dashed var(--border2);border-radius:var(--radius);padding:44px 20px;text-align:center;cursor:pointer;transition:all .2s;background:var(--bg2)}
.drop-zone:hover,.drop-zone.dragover{border-color:var(--green);background:var(--green-bg)}
.drop-icon{font-size:36px;margin-bottom:10px}
.drop-title{font-size:15px;font-weight:500;margin-bottom:5px}
.drop-sub{font-size:13px;color:var(--text2)}
.import-preview{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:20px;margin-top:18px}
.import-stats{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:14px 0}
.istat{background:var(--bg3);border-radius:var(--radius-sm);padding:12px;text-align:center}
.istat-val{font-size:20px;font-weight:600;font-family:'DM Mono',monospace}
.istat-label{font-size:11px;color:var(--text2);margin-top:3px}
.import-table-wrap{max-height:260px;overflow-y:auto;border:1px solid var(--border);border-radius:var(--radius-sm);margin-top:14px}
.import-table{width:100%;border-collapse:collapse;font-size:12px}
.import-table th{padding:8px 12px;background:var(--bg3);color:var(--text2);font-weight:500;text-align:left;position:sticky;top:0}
.import-table td{padding:8px 12px;border-bottom:1px solid var(--border)}
.import-table tr:last-child td{border-bottom:none}
.import-warn{background:rgba(251,191,36,.1);border:1px solid rgba(251,191,36,.3);border-radius:var(--radius-sm);padding:11px 14px;font-size:13px;color:var(--amber);margin-top:10px}
.cloud-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:24px;margin-bottom:20px}
.cloud-step{display:flex;gap:14px;margin-bottom:18px;align-items:flex-start}
.cloud-step-num{width:26px;height:26px;border-radius:50%;background:var(--green-bg);border:1px solid var(--green);color:var(--green);font-size:12px;font-weight:600;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:1px}
.cloud-step-body{flex:1}
.cloud-step-title{font-size:13.5px;font-weight:500;margin-bottom:4px}
.cloud-step-desc{font-size:12.5px;color:var(--text2);line-height:1.5}
.cloud-step-desc a{color:var(--blue);text-decoration:none}
.sync-status{display:inline-flex;align-items:center;gap:6px;font-size:12px;padding:4px 10px;border-radius:20px}
.sync-dot{width:7px;height:7px;border-radius:50%}
.sync-ok{background:var(--green-bg);color:var(--green)}
.sync-dot.ok{background:var(--green)}
.sync-err{background:var(--red-bg);color:var(--red)}
.sync-dot.err{background:var(--red)}
.sync-local{background:var(--bg3);color:var(--text2)}
.sync-dot.local{background:var(--text3)}
.section{display:none}
.section.active{display:block}
.empty{text-align:center;padding:36px 20px;color:var(--text3);font-size:13px}
.empty-icon{font-size:28px;margin-bottom:8px}
/* Insights */
.insights-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px}
.insight-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:18px}
.insight-label{font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
.insight-value{font-size:20px;font-weight:600;font-family:'DM Mono',monospace}
.insight-sub{font-size:11px;color:var(--text2);margin-top:4px}
/* Patrimônio */
.pat-input-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:20px}
.pat-field{display:flex;flex-direction:column;gap:5px}
.pat-field label{font-size:11.5px;color:var(--text2)}
.pat-input{width:100%;padding:9px 11px;background:var(--bg3);border:1px solid var(--border2);border-radius:var(--radius-sm);color:var(--text);font-family:'DM Sans',sans-serif;font-size:13px;outline:none;transition:border .15s}
.pat-input:focus{border-color:var(--green)}
.pat-input-hint{font-size:10.5px;color:var(--text3);margin-top:2px}
.pat-api-badge{display:inline-flex;align-items:center;gap:5px;font-size:11px;padding:3px 9px;border-radius:20px;margin-left:8px}
.pat-api-ok{background:var(--green-bg);color:var(--green)}
.pat-api-err{background:var(--red-bg);color:var(--red)}
.pat-api-load{background:var(--bg3);color:var(--text3)}
.pat-milestone{display:inline-flex;align-items:center;gap:6px;padding:5px 12px;border-radius:8px;font-size:12px;background:var(--bg3);border:1px solid var(--border)}
.pat-milestone.reached{background:var(--green-bg);border-color:var(--green);color:var(--green)}
.pat-table-wrap{max-height:420px;overflow-y:auto}
.pat-table{width:100%;border-collapse:collapse;font-size:12px}
.pat-table th{padding:9px 14px;background:var(--bg3);color:var(--text2);font-weight:500;text-align:right;position:sticky;top:0}
.pat-table th:first-child{text-align:left}
.pat-table td{padding:9px 14px;border-bottom:1px solid var(--border);text-align:right;font-family:'DM Mono',monospace}
.pat-table td:first-child{text-align:left;font-family:'DM Sans',sans-serif;font-weight:500;color:var(--text)}
.pat-table tr:last-child td{border-bottom:none}
.pat-table tr.highlight-year td{background:rgba(34,211,160,0.06)}
.pat-excluded-cats{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}
.cmp-badge{display:inline-flex;align-items:center;gap:4px;font-size:11px;padding:2px 8px;border-radius:20px;font-weight:500}
.cmp-up{background:rgba(34,211,160,0.15);color:var(--green)}
.cmp-down{background:rgba(245,101,101,0.15);color:var(--red)}
.cmp-flat{background:var(--bg3);color:var(--text3)}
.cmp-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px}
.cmp-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);padding:18px}
.cmp-label{font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}
.cmp-row{display:flex;justify-content:space-between;align-items:center;margin-bottom:5px}
.cmp-year-label{font-size:11px;color:var(--text2)}
.cmp-value{font-family:'DM Mono',monospace;font-size:13px;font-weight:500}
.cmp-divider{height:1px;background:var(--border);margin:8px 0}
.rank-row{display:flex;align-items:center;gap:10px;padding:10px 0;border-bottom:1px solid var(--border)}
.rank-row:last-child{border-bottom:none}
.rank-num{width:22px;height:22px;border-radius:50%;background:var(--bg3);color:var(--text2);font-size:11px;font-weight:600;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.rank-bar-wrap{flex:1;height:4px;background:var(--bg3);border-radius:2px;overflow:hidden}
.rank-bar{height:100%;border-radius:2px}
.mobile-bar{display:none;position:fixed;bottom:0;left:0;right:0;background:var(--bg2);border-top:1px solid var(--border);z-index:100;padding:6px 0 calc(6px + env(safe-area-inset-bottom))}
/* Lock screen */
.lock-screen{position:fixed;inset:0;background:var(--bg);z-index:9999;display:flex;align-items:center;justify-content:center}
.lock-card{background:var(--bg2);border:1px solid var(--border2);border-radius:var(--radius);padding:40px 36px;width:360px;max-width:95vw;text-align:center}
.lock-logo{width:56px;height:56px;background:linear-gradient(135deg,var(--green),#0ea5e9);border-radius:16px;display:flex;align-items:center;justify-content:center;font-size:26px;margin:0 auto 20px}
.lock-title{font-size:20px;font-weight:600;margin-bottom:6px}
.lock-sub{font-size:13px;color:var(--text2);margin-bottom:28px}
.lock-input{width:100%;padding:12px 16px;background:var(--bg3);border:1px solid var(--border2);border-radius:var(--radius-sm);color:var(--text);font-family:'DM Sans',sans-serif;font-size:15px;outline:none;text-align:center;letter-spacing:4px;transition:border .2s}
.lock-input:focus{border-color:var(--green)}
.lock-input.err{border-color:var(--red);animation:shake .3s}
@keyframes shake{0%,100%{transform:translateX(0)}25%{transform:translateX(-6px)}75%{transform:translateX(6px)}}
.lock-btn{width:100%;margin-top:14px;padding:11px;border-radius:var(--radius-sm);background:var(--green);color:#0f1117;font-size:14px;font-weight:600;border:none;cursor:pointer;font-family:'DM Sans',sans-serif;transition:background .15s}
.lock-btn:hover{background:#1ab88c}
.lock-err{color:var(--red);font-size:12px;margin-top:10px;min-height:18px}
.mobile-bar-inner{display:flex;justify-content:space-around}
.mob-btn{display:flex;flex-direction:column;align-items:center;gap:3px;padding:6px 12px;cursor:pointer;color:var(--text3);font-size:10px;border-radius:8px;transition:color .15s;border:none;background:none}
.mob-btn.active{color:var(--green)}
.mob-btn .mob-icon{font-size:18px}
@media(max-width:768px){
  :root{--sidebar:0px}
  .sidebar{transform:translateX(-220px);width:220px}
  .sidebar.open{transform:translateX(0)}
  .main{margin-left:0;padding:16px 16px 80px}
  .cards-grid{grid-template-columns:repeat(2,1fr);gap:10px}
  .content-grid{grid-template-columns:1fr}
  .cat-manager{grid-template-columns:1fr}
  .import-stats{grid-template-columns:repeat(2,1fr)}
  .form-row-2{grid-template-columns:1fr}
  .mobile-bar{display:block}
  .insights-grid{grid-template-columns:repeat(2,1fr)}
}
@media(min-width:769px) and (max-width:1024px){
  :root{--sidebar:60px}
  .sidebar .nav-item span,.sidebar .logo span{display:none}
  .cards-grid{grid-template-columns:repeat(2,1fr)}
}
</style>
</head>
<body>
<!-- Lock screen -->
<div id="lock-screen" class="lock-screen">
  <div class="lock-card">
    <div class="lock-logo">💰</div>
    <div class="lock-title">FinTrack</div>
    <div class="lock-sub" id="lock-sub">Quem está entrando?</div>
    <div class="profile-picker" id="profile-picker" style="display:flex;gap:10px;margin-bottom:6px">
      <div class="profile-opt" id="profile-opt-thiago" onclick="pickProfile('thiago')" style="flex:1;padding:16px 10px;border-radius:var(--radius-sm);border:1px solid var(--border2);background:var(--bg3);cursor:pointer;text-align:center;transition:all .15s">
        <div style="font-size:24px;margin-bottom:6px">👤</div><div style="font-size:13px;font-weight:500">Thiago</div>
      </div>
      <div class="profile-opt" id="profile-opt-luiza" onclick="pickProfile('luiza')" style="flex:1;padding:16px 10px;border-radius:var(--radius-sm);border:1px solid var(--border2);background:var(--bg3);cursor:pointer;text-align:center;transition:all .15s">
        <div style="font-size:24px;margin-bottom:6px">👤</div><div style="font-size:13px;font-weight:500">Luiza</div>
      </div>
    </div>
    <div id="lock-pass-row" style="display:none">
      <input class="lock-input" id="lock-input" type="password" placeholder="••••••••" maxlength="20"
        onkeydown="if(event.key==='Enter')checkPassword()" style="margin-top:18px">
      <button class="lock-btn" onclick="checkPassword()">Entrar</button>
      <div style="margin-top:10px"><a href="#" onclick="backToProfilePick(event)" style="font-size:12px;color:var(--text2);text-decoration:none">← escolher outro perfil</a></div>
    </div>
    <div class="lock-err" id="lock-err"></div>
  </div>
</div>

<div class="app" id="app-root" style="display:none">

<aside class="sidebar" id="sidebar">
  <div class="logo"><div class="logo-icon">💰</div><span>FinTrack</span></div>

  <div style="padding:0 14px 14px;border-bottom:1px solid var(--border);margin-bottom:10px">
    <div style="font-size:10.5px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">Logado como</div>
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">
      <span id="profile-badge" style="font-size:13px;font-weight:500">—</span>
      <a href="#" onclick="logoutProfile();return false" style="font-size:11px;color:var(--blue);text-decoration:none">Trocar</a>
    </div>
    <div style="font-size:10.5px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">Visualizando</div>
    <div style="display:flex;gap:5px" id="view-switcher">
      <div class="view-chip" id="view-thiago" onclick="setViewOwner('thiago')" style="flex:1;text-align:center;padding:6px 4px;border-radius:7px;font-size:11px;cursor:pointer;border:1px solid var(--border2);background:var(--bg3);color:var(--text2)">Thiago</div>
      <div class="view-chip" id="view-luiza" onclick="setViewOwner('luiza')" style="flex:1;text-align:center;padding:6px 4px;border-radius:7px;font-size:11px;cursor:pointer;border:1px solid var(--border2);background:var(--bg3);color:var(--text2)">Luiza</div>
      <div class="view-chip" id="view-casal" onclick="setViewOwner('casal')" style="flex:1;text-align:center;padding:6px 4px;border-radius:7px;font-size:11px;cursor:pointer;border:1px solid var(--border2);background:var(--bg3);color:var(--text2)">Casal</div>
    </div>
  </div>

  <nav class="sidebar-nav">
    <div class="nav-item active" data-section="dashboard" onclick="showSection('dashboard')"><span class="nav-icon">📊</span><span>Dashboard</span></div>
    <div class="nav-item" data-section="transacoes" onclick="showSection('transacoes')"><span class="nav-icon">↕️</span><span>Transações</span></div>
    <div class="nav-item" data-section="evolucao" onclick="showSection('evolucao')"><span class="nav-icon">📈</span><span>Evolução</span></div>
    <div class="nav-item" data-section="comparativo" onclick="showSection('comparativo')"><span class="nav-icon">🔄</span><span>Comparativo</span></div>
    <div class="nav-item" data-section="insights" onclick="showSection('insights')"><span class="nav-icon">🔍</span><span>Insights</span></div>
    <div class="nav-item" data-section="patrimonio" onclick="showSection('patrimonio')"><span class="nav-icon">🏦</span><span>Patrimônio</span></div>
    <div class="nav-item" data-section="categorias" onclick="showSection('categorias')"><span class="nav-icon">🏷️</span><span>Categorias</span></div>
    <div class="nav-item" data-section="anual" onclick="showSection('anual')"><span class="nav-icon">📅</span><span>Visão Anual</span></div>
    <div class="nav-item" data-section="importar" onclick="showSection('importar')"><span class="nav-icon">📥</span><span>Importar Excel</span></div>
    <div class="nav-item" data-section="exportar" onclick="showSection('exportar')"><span class="nav-icon">📤</span><span>Exportar Excel</span></div>
    <div class="nav-item" data-section="gcat" onclick="showSection('gcat')"><span class="nav-icon">⚙️</span><span>Ger. Categorias</span></div>
    <div class="nav-item" data-section="cloud" onclick="showSection('cloud')"><span class="nav-icon">☁️</span><span>Nuvem</span></div>
    <div class="nav-spacer"></div>
    <div class="nav-item" data-section="reset" onclick="showSection('reset')" style="color:var(--red);border-top:1px solid var(--border)"><span class="nav-icon">🗑️</span><span>Zerar dados</span></div>
  </nav>
</aside>

<main class="main">

<!-- DASHBOARD -->
<section id="sec-dashboard" class="section active">
  <div class="page-header">
    <div><div class="page-title">Dashboard</div><div class="page-sub" id="dash-period-label">—</div></div>
    <div class="hdr-actions">
      <select class="period-select" id="dash-month" onchange="updateDashboard()"></select>
      <select class="period-select" id="dash-year" onchange="updateDashboard()"></select>
      <button class="btn btn-primary" onclick="openModal()">+ <span>Nova transação</span></button>
    </div>
  </div>
  <div class="cards-grid">
    <div class="card"><div class="card-label">Receitas</div><div class="card-value" id="c-income" style="color:var(--green)">R$ 0</div><div class="card-sub" id="c-income-count">—</div></div>
    <div class="card"><div class="card-label">Despesas</div><div class="card-value" id="c-expense" style="color:var(--red)">R$ 0</div><div class="card-sub" id="c-expense-count">—</div></div>
    <div class="card"><div class="card-label">Saldo do mês</div><div class="card-value" id="c-balance">R$ 0</div><div class="card-sub" id="c-balance-pct">—</div></div>
    <div class="card"><div class="card-label">Maior gasto</div><div class="card-value" id="c-topcat" style="font-size:14px;padding-top:4px;color:var(--amber)">—</div><div class="card-sub" id="c-topcat-amt">—</div></div>
  </div>
  <div class="chart-card" style="margin-bottom:18px">
    <div class="chart-header"><div class="chart-title">⚖️ Balanço do casal — quem pagou os gastos compartilhados</div></div>
    <div id="balance-card-body" style="font-size:13px;color:var(--text2)">—</div>
  </div>
  <div class="content-grid">
    <div class="chart-card">
      <div class="chart-header">
        <div class="chart-title">Receitas vs Despesas — últimos 6 meses</div>
        <div class="legend"><span class="legend-item"><span class="legend-dot" style="background:var(--green)"></span>Receitas</span><span class="legend-item"><span class="legend-dot" style="background:var(--red)"></span>Despesas</span></div>
      </div>
      <div style="position:relative;height:220px"><canvas id="chart-monthly"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Gastos por categoria</div></div>
      <div style="position:relative;height:170px"><canvas id="chart-pie"></canvas></div>
      <div id="pie-legend" class="legend" style="margin-top:14px;flex-direction:column;gap:7px"></div>
    </div>
  </div>
  <div class="transactions-card" style="margin-top:18px">
    <div class="tx-header"><div class="tx-title">Últimas transações</div><button class="btn" onclick="showSection('transacoes')">Ver todas →</button></div>
    <div class="tx-list" id="recent-list"></div>
  </div>
</section>

<!-- TRANSAÇÕES -->
<section id="sec-transacoes" class="section">
  <div class="page-header">
    <div><div class="page-title">Transações</div><div class="page-sub" id="tx-count-label">—</div></div>
    <button class="btn btn-primary" onclick="openModal()">+ <span>Nova transação</span></button>
  </div>

  <!-- Filtros visuais -->
  <div class="filter-section">
    <div class="filter-section-title">Período</div>
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px">
      <select class="period-select" id="tx-month" onchange="renderTxList()"></select>
      <select class="period-select" id="tx-year" onchange="renderTxList()"></select>
    </div>
    <div class="filter-section-title">Tipo</div>
    <div class="chip-group" style="margin-bottom:14px">
      <div class="filter-chip active" id="chip-type-all" onclick="setTxTypeFilter('all')">Todos</div>
      <div class="filter-chip" id="chip-type-income" onclick="setTxTypeFilter('income')">⬆ Receitas</div>
      <div class="filter-chip" id="chip-type-expense" onclick="setTxTypeFilter('expense')">⬇ Despesas</div>
    </div>
    <div class="filter-section-title">Categoria <span style="color:var(--text3);font-weight:400;font-size:11px">(clique para filtrar, clique novamente para remover)</span></div>
    <div class="chip-group" id="cat-filter-chips"></div>
  </div>

  <div class="transactions-card">
    <div class="tx-header">
      <div class="tx-title" id="tx-list-title">Todas as transações</div>
      <button class="btn btn-danger" id="btn-delete-filtered" onclick="deleteFiltered()" style="display:none;font-size:12px">🗑 Excluir selecionadas</button>
    </div>
    <div class="tx-list" id="tx-list" style="max-height:none"></div>
  </div>
</section>

<!-- EVOLUÇÃO MENSAL -->
<section id="sec-evolucao" class="section">
  <div class="page-header">
    <div><div class="page-title">Evolução Mensal</div><div class="page-sub">Acompanhe receitas e despesas mês a mês</div></div>
    <div class="hdr-actions">
      <select class="period-select" id="ev-year" onchange="renderEvolucao()"></select>
    </div>
  </div>
  <div class="cards-grid" style="grid-template-columns:repeat(3,1fr)">
    <div class="card"><div class="card-label">Melhor mês (receita)</div><div class="card-value" id="ev-best-income" style="color:var(--green);font-size:16px">—</div><div class="card-sub" id="ev-best-income-val">—</div></div>
    <div class="card"><div class="card-label">Pior mês (despesa)</div><div class="card-value" id="ev-worst-expense" style="color:var(--red);font-size:16px">—</div><div class="card-sub" id="ev-worst-expense-val">—</div></div>
    <div class="card"><div class="card-label">Saldo acumulado</div><div class="card-value" id="ev-accumulated">R$ 0</div><div class="card-sub" id="ev-accumulated-sub">—</div></div>
  </div>
  <div class="chart-card" style="margin-bottom:18px">
    <div class="chart-header">
      <div class="chart-title">Receitas, Despesas e Saldo mensais</div>
      <div class="legend">
        <span class="legend-item"><span class="legend-dot" style="background:var(--green)"></span>Receitas</span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--red)"></span>Despesas</span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--blue)"></span>Saldo</span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--amber);border-radius:50%"></span>Saldo acum.</span>
      </div>
    </div>
    <div style="position:relative;height:300px"><canvas id="chart-evolucao"></canvas></div>
  </div>
  <div class="content-grid">
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Saldo acumulado no ano</div></div>
      <div style="position:relative;height:220px"><canvas id="chart-acumulado"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Médias mensais</div></div>
      <div id="ev-medias" style="padding:4px 0"></div>
    </div>
  </div>
</section>

<!-- COMPARATIVO ANO ANTERIOR -->
<section id="sec-comparativo" class="section">
  <div class="page-header">
    <div><div class="page-title">Comparativo</div><div class="page-sub" id="cmp-subtitle">Período atual vs ano anterior</div></div>
    <div class="hdr-actions">
      <select class="period-select" id="cmp-mode" onchange="renderComparativo()">
        <option value="ytd">Acumulado no ano (Jan–mês atual)</option>
        <option value="month">Mês específico</option>
        <option value="full">Ano completo</option>
      </select>
      <select class="period-select" id="cmp-month" onchange="renderComparativo()" style="display:none"></select>
      <select class="period-select" id="cmp-year" onchange="renderComparativo()"></select>
    </div>
  </div>

  <!-- Cards resumo -->
  <div class="cmp-grid">
    <div class="cmp-card">
      <div class="cmp-label">Receitas</div>
      <div class="cmp-row"><span class="cmp-year-label" id="cmp-year-cur-label">2025</span><span class="cmp-value" id="cmp-inc-cur" style="color:var(--green)">—</span></div>
      <div class="cmp-row"><span class="cmp-year-label" id="cmp-year-prev-label">2024</span><span class="cmp-value" id="cmp-inc-prev" style="color:var(--text3)">—</span></div>
      <div class="cmp-divider"></div>
      <div style="display:flex;align-items:center;justify-content:space-between">
        <span style="font-size:12px;color:var(--text2)">Variação</span>
        <span id="cmp-inc-badge">—</span>
      </div>
    </div>
    <div class="cmp-card">
      <div class="cmp-label">Despesas</div>
      <div class="cmp-row"><span class="cmp-year-label" id="cmp-year-cur-label2">2025</span><span class="cmp-value" id="cmp-exp-cur" style="color:var(--red)">—</span></div>
      <div class="cmp-row"><span class="cmp-year-label" id="cmp-year-prev-label2">2024</span><span class="cmp-value" id="cmp-exp-prev" style="color:var(--text3)">—</span></div>
      <div class="cmp-divider"></div>
      <div style="display:flex;align-items:center;justify-content:space-between">
        <span style="font-size:12px;color:var(--text2)">Variação</span>
        <span id="cmp-exp-badge">—</span>
      </div>
    </div>
    <div class="cmp-card">
      <div class="cmp-label">Saldo</div>
      <div class="cmp-row"><span class="cmp-year-label" id="cmp-year-cur-label3">2025</span><span class="cmp-value" id="cmp-bal-cur">—</span></div>
      <div class="cmp-row"><span class="cmp-year-label" id="cmp-year-prev-label3">2024</span><span class="cmp-value" id="cmp-bal-prev" style="color:var(--text3)">—</span></div>
      <div class="cmp-divider"></div>
      <div style="display:flex;align-items:center;justify-content:space-between">
        <span style="font-size:12px;color:var(--text2)">Variação</span>
        <span id="cmp-bal-badge">—</span>
      </div>
    </div>
  </div>

  <!-- Gráfico comparativo mês a mês -->
  <div class="chart-card" style="margin-bottom:18px">
    <div class="chart-header">
      <div class="chart-title" id="cmp-chart-title">Despesas mês a mês — ano atual vs anterior</div>
      <div class="legend">
        <span class="legend-item"><span class="legend-dot" style="background:var(--red)"></span><span id="cmp-legend-cur">2025</span></span>
        <span class="legend-item"><span class="legend-dot" style="background:rgba(245,101,101,0.3)"></span><span id="cmp-legend-prev">2024</span></span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--green)"></span><span id="cmp-legend-inc-cur">Receita 2025</span></span>
        <span class="legend-item"><span class="legend-dot" style="background:rgba(34,211,160,0.3)"></span><span id="cmp-legend-inc-prev">Receita 2024</span></span>
      </div>
    </div>
    <div style="position:relative;height:300px"><canvas id="chart-comparativo"></canvas></div>
  </div>

  <!-- Comparativo por categoria -->
  <div class="content-grid">
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Despesas por categoria — comparativo</div></div>
      <div id="cmp-cat-list" class="cat-list"></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Maiores variações</div></div>
      <div id="cmp-variations"></div>
    </div>
  </div>
</section>

<!-- INSIGHTS -->
<section id="sec-insights" class="section">
  <div class="page-header">
    <div><div class="page-title">Insights</div><div class="page-sub">Análise detalhada do seu histórico financeiro</div></div>
    <select class="period-select" id="ins-year" onchange="renderInsights()"></select>
  </div>
  <div class="insights-grid">
    <div class="insight-card"><div class="insight-label">Maior receita no ano</div><div class="insight-value" id="ins-top-income" style="color:var(--green)">—</div><div class="insight-sub" id="ins-top-income-sub">—</div></div>
    <div class="insight-card"><div class="insight-label">Maior despesa no ano</div><div class="insight-value" id="ins-top-expense" style="color:var(--red)">—</div><div class="insight-sub" id="ins-top-expense-sub">—</div></div>
    <div class="insight-card"><div class="insight-label">Taxa média de poupança</div><div class="insight-value" id="ins-save-rate">—%</div><div class="insight-sub" id="ins-save-sub">—</div></div>
  </div>
  <div class="content-grid" style="margin-bottom:20px">
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Ranking de meses — Despesas</div></div>
      <div id="ins-rank-expense"></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Ranking de meses — Receitas</div></div>
      <div id="ins-rank-income"></div>
    </div>
  </div>
  <div class="content-grid">
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">% de gasto por categoria</div></div>
      <div style="position:relative;height:260px"><canvas id="chart-ins-cat"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Detalhamento por categoria</div></div>
      <div id="ins-cat-detail"></div>
    </div>
  </div>
</section>

<!-- CATEGORIAS -->
<section id="sec-categorias" class="section">
  <div class="page-header">
    <div><div class="page-title">Categorias</div><div class="page-sub">Análise por categoria</div></div>
    <div class="hdr-actions">
      <select class="period-select" id="cat-month" onchange="renderCategorias()">
        <option value="-1">Todos os meses</option>
      </select>
      <select class="period-select" id="cat-year" onchange="renderCategorias()"></select>
    </div>
  </div>
  <div class="content-grid">
    <div>
      <div class="chart-card" style="margin-bottom:18px"><div class="chart-header"><div class="chart-title">Despesas por categoria</div></div><div class="cat-list" id="cat-expense-bars"></div></div>
      <div class="chart-card"><div class="chart-header"><div class="chart-title">Receitas por categoria</div></div><div class="cat-list" id="cat-income-bars"></div></div>
    </div>
    <div class="chart-card">
      <div class="chart-header"><div class="chart-title">Distribuição de despesas</div></div>
      <div style="position:relative;height:260px"><canvas id="chart-cat-pie"></canvas></div>
      <div id="cat-pie-legend" class="legend" style="margin-top:14px;flex-direction:column;gap:7px"></div>
    </div>
  </div>
</section>

<!-- VISÃO ANUAL -->
<section id="sec-anual" class="section">
  <div class="page-header">
    <div><div class="page-title">Visão Anual</div><div class="page-sub">Resumo e análise por categoria</div></div>
    <select class="period-select" id="anual-year" onchange="renderAnual()"></select>
  </div>
  <div class="cards-grid" style="grid-template-columns:repeat(3,1fr)">
    <div class="card"><div class="card-label">Total Receitas</div><div class="card-value" id="anual-income" style="color:var(--green)">R$ 0</div></div>
    <div class="card"><div class="card-label">Total Despesas</div><div class="card-value" id="anual-expense" style="color:var(--red)">R$ 0</div></div>
    <div class="card"><div class="card-label">Saldo Anual</div><div class="card-value" id="anual-balance">R$ 0</div></div>
  </div>
  <div class="chart-card" style="margin-bottom:18px">
    <div class="chart-header">
      <div class="chart-title">Mensal — Receitas vs Despesas</div>
      <div class="legend"><span class="legend-item"><span class="legend-dot" style="background:var(--green)"></span>Receitas</span><span class="legend-item"><span class="legend-dot" style="background:var(--red)"></span>Despesas</span><span class="legend-item"><span class="legend-dot" style="background:var(--blue)"></span>Saldo</span></div>
    </div>
    <div style="position:relative;height:260px"><canvas id="chart-anual"></canvas></div>
  </div>
  <div class="chart-card" style="margin-bottom:18px">
    <div class="chart-header">
      <div class="chart-title">Evolução anual por categoria</div>
      <div class="hdr-actions" style="gap:8px">
        <select class="period-select" id="cat-anual-type" onchange="populateCatAnualSelect();renderCatAnual()" style="font-size:12px"><option value="expense">Despesas</option><option value="income">Receitas</option></select>
        <select class="period-select" id="cat-anual-sel" onchange="renderCatAnual()" style="font-size:12px"></select>
      </div>
    </div>
    <div style="position:relative;height:260px"><canvas id="chart-cat-anual"></canvas></div>
    <div id="cat-anual-summary" style="display:flex;gap:12px;flex-wrap:wrap;margin-top:14px;font-size:12px"></div>
  </div>
  <div class="transactions-card">
    <div class="tx-header"><div class="tx-title">Resumo mensal</div></div>
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead><tr style="border-bottom:1px solid var(--border)">
          <th style="padding:11px 18px;text-align:left;color:var(--text2);font-weight:500">Mês</th>
          <th style="padding:11px 18px;text-align:right;color:var(--text2);font-weight:500">Receitas</th>
          <th style="padding:11px 18px;text-align:right;color:var(--text2);font-weight:500">Despesas</th>
          <th style="padding:11px 18px;text-align:right;color:var(--text2);font-weight:500">Saldo</th>
          <th style="padding:11px 18px;text-align:right;color:var(--text2);font-weight:500">Qtd.</th>
        </tr></thead>
        <tbody id="anual-tbody"></tbody>
      </table>
    </div>
  </div>
</section>

<!-- IMPORTAR -->
<section id="sec-importar" class="section">
  <div class="page-header"><div><div class="page-title">Importar Excel</div><div class="page-sub">Compatível com a sua planilha</div></div></div>
  <div class="drop-zone" id="drop-zone" onclick="document.getElementById('file-input').click()" ondragover="dragOver(event)" ondragleave="dragLeave()" ondrop="dropFile(event)">
    <input type="file" id="file-input" accept=".xlsx,.xls" style="display:none" onchange="handleFile(this.files[0])">
    <div class="drop-icon">📂</div>
    <div class="drop-title">Arraste o arquivo Excel aqui</div>
    <div class="drop-sub">ou clique para selecionar • .xlsx / .xls</div>
    <div class="drop-sub" style="margin-top:6px;font-size:11px;opacity:.6">Formato: Dia | Gasto/Recebido | Departamento | Valor | Descrição</div>
  </div>
  <div id="import-preview" style="display:none">
    <div class="import-preview">
      <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px">
        <div><div style="font-size:14px;font-weight:500" id="import-filename">—</div><div style="font-size:12px;color:var(--text2);margin-top:2px" id="import-fileinfo">—</div></div>
        <div style="display:flex;gap:8px;align-items:center">
          <select class="form-select" id="imp-owner" style="width:auto"><option value="thiago">Importar como: Thiago</option><option value="luiza">Importar como: Luiza</option></select>
          <button class="btn" onclick="resetImport()">✕ Cancelar</button><button class="btn btn-primary" id="btn-confirm-import" onclick="confirmImport()">✓ Importar</button>
        </div>
      </div>
      <div class="import-stats">
        <div class="istat"><div class="istat-val" id="imp-total">0</div><div class="istat-label">Lidas</div></div>
        <div class="istat"><div class="istat-val" id="imp-income" style="color:var(--green)">0</div><div class="istat-label">Receitas</div></div>
        <div class="istat"><div class="istat-val" id="imp-expense" style="color:var(--red)">0</div><div class="istat-label">Despesas</div></div>
        <div class="istat"><div class="istat-val" id="imp-skip" style="color:var(--amber)">0</div><div class="istat-label">Ignoradas</div></div>
      </div>
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;flex-wrap:wrap;gap:8px">
        <div style="font-size:13px;font-weight:500">Prévia (20 primeiras)</div>
        <label style="display:flex;align-items:center;gap:6px;font-size:12px;color:var(--text2);cursor:pointer"><input type="checkbox" id="imp-skip-dup" checked style="accent-color:var(--green)"> Ignorar duplicatas</label>
      </div>
      <div class="import-table-wrap"><table class="import-table"><thead><tr><th>Data</th><th>Tipo</th><th>Categoria</th><th>Valor</th><th>Descrição</th></tr></thead><tbody id="import-tbody"></tbody></table></div>
      <div id="import-warn" class="import-warn" style="display:none"></div>
      <div style="margin-top:14px"><div style="font-size:12px;color:var(--text2);margin-bottom:6px">Novas categorias:</div><div id="import-new-cats" style="display:flex;gap:6px;flex-wrap:wrap"></div></div>
    </div>
  </div>
  <div class="chart-card" style="margin-top:18px"><div class="chart-header"><div class="chart-title">Histórico de importações</div></div><div id="import-history-list"><div class="empty"><div class="empty-icon">📋</div>Nenhuma importação ainda</div></div></div>
</section>

<!-- EXPORTAR -->
<section id="sec-exportar" class="section">
  <div class="page-header"><div><div class="page-title">Exportar Excel</div><div class="page-sub">Baixe seus dados em formato .xlsx</div></div></div>
  <div class="chart-card" style="margin-bottom:18px">
    <div style="font-size:14px;font-weight:500;margin-bottom:16px">Filtros de exportação</div>
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:20px">
      <div><div class="form-label">De quem</div><select class="form-select" id="exp-owner" onchange="renderExportSection()"><option value="ambos">Ambos</option><option value="thiago">Thiago</option><option value="luiza">Luiza</option></select></div>
      <div><div class="form-label">Mês</div><select class="form-select" id="exp-month"><option value="-1">Todos os meses</option></select></div>
      <div><div class="form-label">Ano</div><select class="form-select" id="exp-year"></select></div>
      <div><div class="form-label">Tipo</div><select class="form-select" id="exp-type"><option value="all">Todos</option><option value="expense">Apenas Despesas</option><option value="income">Apenas Receitas</option></select></div>
      <div><div class="form-label">Categoria</div><select class="form-select" id="exp-cat"><option value="all">Todas</option></select></div>
    </div>
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <button class="btn btn-primary" onclick="exportExcel()" style="font-size:13px;padding:9px 20px">📥 Baixar Excel</button>
      <span id="exp-preview-count" style="font-size:12px;color:var(--text2)">—</span>
    </div>
  </div>
  <div class="chart-card"><div style="font-size:14px;font-weight:500;margin-bottom:14px">Prévia</div><div class="import-table-wrap" style="max-height:380px"><table class="import-table"><thead><tr><th>Data</th><th>Tipo</th><th>Dono</th><th>Categoria</th><th>Valor</th><th>Descrição</th></tr></thead><tbody id="exp-preview-tbody"></tbody></table></div></div>
</section>

<!-- PATRIMÔNIO -->
<section id="sec-patrimonio" class="section">
  <div class="page-header">
    <div>
      <div class="page-title">Evolução Patrimonial</div>
      <div class="page-sub">Projeção baseada no seu histórico + SELIC + IPCA projetados pelo BCB</div>
    </div>
    <button class="btn btn-primary" onclick="calcularPatrimonio()">▶ Calcular</button>
  </div>

  <!-- Parâmetros -->
  <div class="chart-card" style="margin-bottom:18px">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:8px">
      <div style="font-size:14px;font-weight:500">Parâmetros da projeção</div>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
        <span id="pat-selic-badge" class="pat-api-badge pat-api-load">⏳ Buscando SELIC...</span>
        <span id="pat-ipca-badge" class="pat-api-badge pat-api-load">⏳ Buscando IPCA...</span>
      </div>
    </div>

    <div class="pat-input-grid">
      <div class="pat-field">
        <label>💰 Saldo atual aplicado (R$)</label>
        <input class="pat-input" id="pat-saldo" type="number" step="0.01" placeholder="Ex: 50000">
        <span class="pat-input-hint">Valor total que você tem aplicado hoje</span>
      </div>
      <div class="pat-field">
        <label>📅 Horizonte de projeção (anos)</label>
        <input class="pat-input" id="pat-anos" type="number" min="1" max="50" value="10" placeholder="Ex: 10">
        <span class="pat-input-hint">Quantos anos quer projetar</span>
      </div>
      <div class="pat-field">
        <label>📈 SELIC anual projetada (%)</label>
        <input class="pat-input" id="pat-selic" type="number" step="0.01" placeholder="Buscando...">
        <span class="pat-input-hint" id="pat-selic-hint">Buscando Boletim Focus...</span>
      </div>
      <div class="pat-field">
        <label>📊 IPCA anual projetado (%)</label>
        <input class="pat-input" id="pat-ipca" type="number" step="0.01" placeholder="Buscando...">
        <span class="pat-input-hint" id="pat-ipca-hint">Buscando Boletim Focus...</span>
      </div>
      <div class="pat-field">
        <label>⬆ Crescimento anual de receita (%)</label>
        <input class="pat-input" id="pat-cresc-rec" type="number" step="0.1" value="3" placeholder="Ex: 3">
        <span class="pat-input-hint">Ex: aumento de salário esperado</span>
      </div>
      <div class="pat-field">
        <label>⬇ Crescimento anual de despesa (%)</label>
        <input class="pat-input" id="pat-cresc-desp" type="number" step="0.1" value="4" placeholder="Ex: 4">
        <span class="pat-input-hint">Ex: inflação sobre seus gastos</span>
      </div>
      <div class="pat-field">
        <label>➕ Aporte extra mensal (R$)</label>
        <input class="pat-input" id="pat-aporte" type="number" step="0.01" value="0" placeholder="0">
        <span class="pat-input-hint">Depósito adicional todo mês</span>
      </div>
    </div>

    <!-- Receita e despesa base calculadas -->
    <div style="background:var(--bg3);border-radius:var(--radius-sm);padding:14px 16px;margin-top:4px">
      <div style="font-size:12px;color:var(--text2);margin-bottom:10px;font-weight:500">Base calculada do histórico (últimos 12 meses)</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
        <div>
          <div style="font-size:11px;color:var(--text3);margin-bottom:4px">Receita operacional média/mês</div>
          <div style="font-size:16px;font-weight:600;font-family:DM Mono;color:var(--green)" id="pat-rec-base">—</div>
          <div style="font-size:10.5px;color:var(--text3);margin-top:3px">Categorias incluídas: <span id="pat-rec-cats" style="color:var(--text2)">—</span></div>
        </div>
        <div>
          <div style="font-size:11px;color:var(--text3);margin-bottom:4px">Despesa média/mês</div>
          <div style="font-size:16px;font-weight:600;font-family:DM Mono;color:var(--red)" id="pat-desp-base">—</div>
          <div style="font-size:10.5px;color:var(--text3);margin-top:3px">Todas as categorias de despesa</div>
        </div>
      </div>
      <div style="margin-top:10px;padding-top:10px;border-top:1px solid var(--border)">
        <div style="font-size:10.5px;color:var(--text3);margin-bottom:4px">Categorias excluídas da projeção de receita:</div>
        <div class="pat-excluded-cats" id="pat-excluded-cats"></div>
      </div>
    </div>
  </div>

  <!-- Cards resultado -->
  <div class="cards-grid" id="pat-cards" style="display:none">
    <div class="card"><div class="card-label">Patrimônio em 1 ano</div><div class="card-value" id="pat-1y" style="color:var(--green)">—</div><div class="card-sub" id="pat-1y-real">—</div></div>
    <div class="card"><div class="card-label" id="pat-mid-label">Patrimônio em 5 anos</div><div class="card-value" id="pat-mid" style="color:var(--green)">—</div><div class="card-sub" id="pat-mid-real">—</div></div>
    <div class="card"><div class="card-label" id="pat-long-label">Patrimônio em 10 anos</div><div class="card-value" id="pat-long" style="color:var(--green)">—</div><div class="card-sub" id="pat-long-real">—</div></div>
    <div class="card"><div class="card-label">Rendimento total acumulado</div><div class="card-value" id="pat-rend-total" style="color:var(--amber)">—</div><div class="card-sub" id="pat-rend-pct">—</div></div>
  </div>

  <!-- Marcos -->
  <div id="pat-milestones" style="display:none;margin-bottom:18px">
    <div style="font-size:13px;font-weight:500;margin-bottom:10px">🏆 Marcos projetados</div>
    <div style="display:flex;gap:8px;flex-wrap:wrap" id="pat-milestone-list"></div>
  </div>

  <!-- Gráfico -->
  <div class="chart-card" style="margin-bottom:18px" id="pat-chart-card" style="display:none">
    <div class="chart-header">
      <div class="chart-title">Evolução patrimonial projetada</div>
      <div class="legend">
        <span class="legend-item"><span class="legend-dot" style="background:var(--green)"></span>Patrimônio nominal</span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--blue)"></span>Valor real (descontado IPCA)</span>
        <span class="legend-item"><span class="legend-dot" style="background:var(--amber)"></span>Rendimento acumulado</span>
      </div>
    </div>
    <div style="position:relative;height:320px"><canvas id="chart-patrimonio"></canvas></div>
  </div>

  <!-- Tabela anual -->
  <div class="transactions-card" id="pat-table-card" style="display:none">
    <div class="tx-header">
      <div class="tx-title">Projeção ano a ano</div>
      <label style="display:flex;align-items:center;gap:6px;font-size:12px;color:var(--text2);cursor:pointer">
        <input type="checkbox" id="pat-show-monthly" onchange="toggleMonthlyDetail()" style="accent-color:var(--green)">
        Ver detalhamento mensal
      </label>
    </div>
    <div class="pat-table-wrap">
      <table class="pat-table">
        <thead>
          <tr>
            <th style="text-align:left">Período</th>
            <th>Receita</th>
            <th>Despesa</th>
            <th>Fluxo</th>
            <th>Rendimento</th>
            <th>Patrimônio Nominal</th>
            <th>Valor Real</th>
          </tr>
        </thead>
        <tbody id="pat-tbody"></tbody>
      </table>
    </div>
  </div>
</section>

<!-- GER. CATEGORIAS -->
<section id="sec-gcat" class="section">
  <div class="page-header"><div><div class="page-title">Gerenciar Categorias</div><div class="page-sub">Crie e edite categorias personalizadas</div></div></div>
  <div class="cat-manager">
    <div class="cat-section-card"><div class="cat-section-title">⬇ Despesas</div><div id="gcat-expense-list"></div><div class="new-cat-form"><input class="form-input" id="new-exp-name" placeholder="Nome" style="flex:1;min-width:100px"><input type="color" class="color-picker" id="new-exp-color" value="#f56565"><input class="form-input" id="new-exp-icon" placeholder="🏷️" style="width:46px;text-align:center"><button class="btn btn-primary" onclick="addCat('expense')">+</button></div></div>
    <div class="cat-section-card"><div class="cat-section-title">⬆ Receitas</div><div id="gcat-income-list"></div><div class="new-cat-form"><input class="form-input" id="new-inc-name" placeholder="Nome" style="flex:1;min-width:100px"><input type="color" class="color-picker" id="new-inc-color" value="#22d3a0"><input class="form-input" id="new-inc-icon" placeholder="💰" style="width:46px;text-align:center"><button class="btn btn-primary" onclick="addCat('income')">+</button></div></div>
  </div>
</section>

<!-- NUVEM -->
<section id="sec-cloud" class="section">
  <div class="page-header">
    <div><div class="page-title">Sincronização na Nuvem</div><div class="page-sub">Acesse seus dados em qualquer dispositivo</div></div>
    <div id="sync-status-badge" class="sync-status sync-local"><span class="sync-dot local"></span>Armazenamento local</div>
  </div>
  <div class="cloud-card">
    <div style="font-size:15px;font-weight:500;margin-bottom:6px">Como funciona</div>
    <div style="font-size:13px;color:var(--text2);margin-bottom:20px;line-height:1.6">Usa o <strong style="color:var(--text)">Firebase Realtime Database</strong> (gratuito até 1 GB) para sincronizar em tempo real.</div>
    <div>
      <div class="cloud-step"><div class="cloud-step-num">1</div><div class="cloud-step-body"><div class="cloud-step-title">Crie um projeto Firebase</div><div class="cloud-step-desc">Acesse <a href="https://console.firebase.google.com" target="_blank">console.firebase.google.com</a> → "Adicionar projeto".</div></div></div>
      <div class="cloud-step"><div class="cloud-step-num">2</div><div class="cloud-step-body"><div class="cloud-step-title">Ative o Realtime Database</div><div class="cloud-step-desc">Menu lateral → "Realtime Database" → "Criar banco de dados" → modo de teste → Ativar.</div></div></div>
      <div class="cloud-step"><div class="cloud-step-num">3</div><div class="cloud-step-body"><div class="cloud-step-title">Copie a URL e cole abaixo</div><div class="cloud-step-desc">URL parecida com <code style="background:var(--bg3);padding:1px 6px;border-radius:4px;font-size:11px">https://seu-projeto-rtdb.firebaseio.com</code></div></div></div>
    </div>
    <div class="form-row" style="margin-bottom:0">
      <label class="form-label">URL do Firebase</label>
      <div style="display:flex;gap:8px"><input class="form-input" id="firebase-url" placeholder="https://seu-projeto-rtdb.firebaseio.com" style="flex:1"><button class="btn btn-primary" onclick="connectFirebase()">Conectar</button></div>
      <div style="font-size:11px;color:var(--text3);margin-top:5px">Ao conectar, novos lançamentos serão enviados automaticamente para a nuvem.</div>
    </div>
  </div>
  <div class="chart-card" id="cloud-actions" style="display:none">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:16px">
      <div><div style="font-size:14px;font-weight:500">Firebase conectado</div><div style="font-size:12px;color:var(--text2);margin-top:2px" id="cloud-url-display">—</div></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn" onclick="syncFromCloud()">⬇ Baixar da nuvem</button>
        <button class="btn btn-primary" onclick="syncToCloud()">⬆ Enviar tudo</button>
        <button class="btn btn-danger" onclick="disconnectFirebase()">Desconectar</button>
      </div>
    </div>
    <div id="cloud-log" style="font-size:12px;color:var(--text2);background:var(--bg3);border-radius:8px;padding:12px;min-height:48px;line-height:1.7"></div>
  </div>
</section>

<!-- ZERAR DADOS -->
<section id="sec-reset" class="section">
  <div class="page-header"><div><div class="page-title">Zerar Dados</div><div class="page-sub">Apague transações, categorias ou tudo</div></div></div>
  <div style="display:grid;gap:16px;max-width:600px">
    <div class="chart-card" style="border-color:rgba(245,101,101,0.2)">
      <div style="display:flex;align-items:flex-start;gap:14px">
        <div style="font-size:32px">📋</div>
        <div style="flex:1">
          <div style="font-size:15px;font-weight:600;margin-bottom:4px">Apagar todas as transações</div>
          <div style="font-size:13px;color:var(--text2);margin-bottom:14px">Remove todos os lançamentos. As categorias são mantidas.</div>
          <div style="background:var(--bg3);border-radius:8px;padding:10px 14px;font-size:12px;color:var(--text2);margin-bottom:14px" id="reset-tx-count">—</div>
          <button class="btn btn-danger" onclick="resetTransactions()">🗑 Apagar transações</button>
        </div>
      </div>
    </div>
    <div class="chart-card" style="border-color:rgba(251,191,36,0.2)">
      <div style="display:flex;align-items:flex-start;gap:14px">
        <div style="font-size:32px">🏷️</div>
        <div style="flex:1">
          <div style="font-size:15px;font-weight:600;margin-bottom:4px">Remover categorias personalizadas</div>
          <div style="font-size:13px;color:var(--text2);margin-bottom:14px">Remove apenas as categorias criadas ou importadas. As padrão são mantidas.</div>
          <div style="background:var(--bg3);border-radius:8px;padding:10px 14px;font-size:12px;color:var(--text2);margin-bottom:14px" id="reset-cats-count">—</div>
          <button class="btn" onclick="resetCustomCats()" style="background:rgba(251,191,36,0.1);border-color:rgba(251,191,36,0.3);color:var(--amber)">🏷 Remover criadas</button>
        </div>
      </div>
    </div>
    <div class="chart-card">
      <div style="display:flex;align-items:flex-start;gap:14px">
        <div style="font-size:32px">📥</div>
        <div style="flex:1">
          <div style="font-size:15px;font-weight:600;margin-bottom:4px">Limpar histórico de importações</div>
          <div style="font-size:13px;color:var(--text2);margin-bottom:14px">Apaga o registro de arquivos importados. Não remove transações.</div>
          <div style="background:var(--bg3);border-radius:8px;padding:10px 14px;font-size:12px;color:var(--text2);margin-bottom:14px" id="reset-imp-count">—</div>
          <button class="btn" onclick="resetImportHistory()">🧹 Limpar histórico</button>
        </div>
      </div>
    </div>
    <div class="chart-card" style="border:2px solid rgba(245,101,101,0.4);background:rgba(245,101,101,0.04)">
      <div style="display:flex;align-items:flex-start;gap:14px">
        <div style="font-size:32px">⚠️</div>
        <div style="flex:1">
          <div style="font-size:15px;font-weight:600;margin-bottom:4px;color:var(--red)">Apagar tudo</div>
          <div style="font-size:13px;color:var(--text2);margin-bottom:14px">Remove <strong style="color:var(--text)">tudo</strong>. Categorias padrão são restauradas. <strong style="color:var(--red)">Irreversível.</strong></div>
          <label style="display:flex;align-items:center;gap:6px;font-size:12px;color:var(--text2);cursor:pointer;margin-bottom:14px"><input type="checkbox" id="confirm-reset-all" style="accent-color:var(--red)"> Entendo que esta ação não pode ser desfeita</label>
          <button class="btn btn-danger" onclick="resetAll()">⚠️ Apagar tudo</button>
        </div>
      </div>
    </div>
  </div>
</section>

</main>
</div><!-- end app-root -->

<!-- Mobile bar -->
<div class="mobile-bar">
  <div class="mobile-bar-inner">
    <button class="mob-btn active" id="mob-dashboard" onclick="showSection('dashboard')"><span class="mob-icon">📊</span>Início</button>
    <button class="mob-btn" id="mob-transacoes" onclick="showSection('transacoes')"><span class="mob-icon">↕️</span>Trans.</button>
    <button class="mob-btn" id="mob-add" onclick="openModal()"><span class="mob-icon">➕</span>Novo</button>
    <button class="mob-btn" id="mob-evolucao" onclick="showSection('evolucao')"><span class="mob-icon">📈</span>Evolução</button>
    <button class="mob-btn" id="mob-menu" onclick="toggleMobileMenu()"><span class="mob-icon">☰</span>Menu</button>
  </div>
</div>

<!-- Modal -->
<div class="modal-bg" id="modal-bg" onclick="closeModalBg(event)">
  <div class="modal">
    <div class="modal-title" id="modal-title">Nova transação</div>
    <div class="form-row">
      <label class="form-label">Tipo</label>
      <div class="radio-group">
        <div class="radio-btn sel-income" id="rb-income" onclick="setType('income')">⬆ Receita</div>
        <div class="radio-btn" id="rb-expense" onclick="setType('expense')">⬇ Despesa</div>
      </div>
    </div>
    <div class="form-row"><label class="form-label">Descrição</label><input class="form-input" id="f-desc" placeholder="Ex: Salário, Aluguel, Mercado..."></div>
    <div class="form-row-2">
      <div class="form-row" style="margin-bottom:0"><label class="form-label">Valor (R$)</label><input class="form-input" id="f-amount" type="number" step="0.01" placeholder="0,00"><div id="f-amount-hint" style="font-size:11px;color:var(--text3);margin-top:4px;display:none">Valor negativo: usado para abater de "a receber" (débito sobre receita)</div></div>
      <div class="form-row" style="margin-bottom:0"><label class="form-label">Data</label><input class="form-input" id="f-date" type="date"></div>
    </div>
    <div class="form-row" style="margin-top:13px">
      <label class="form-label">Categoria</label>
      <input class="form-input" id="f-cat-search" placeholder="Buscar categoria..." oninput="filterCatOptions()" onclick="showCatDropdown()" autocomplete="off" style="margin-bottom:4px">
      <div id="f-cat-dropdown" style="display:none;background:var(--bg3);border:1px solid var(--border2);border-radius:var(--radius-sm);max-height:200px;overflow-y:auto;position:relative;z-index:10"></div>
      <input type="hidden" id="f-cat">
    </div>
    <!-- Compartilhado -->
    <div class="form-row" style="margin-top:4px" id="f-shared-row">
      <label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13px;color:var(--text2)">
        <input type="checkbox" id="f-shared" onchange="toggleShared()" style="accent-color:var(--green)">
        🤝 Compartilhado com <span id="f-shared-partner-name"></span>
      </label>
    </div>
    <div id="f-shared-split-row" style="display:none;background:var(--bg3);border-radius:var(--radius-sm);padding:12px 14px;margin-bottom:13px">
      <div style="font-size:11.5px;color:var(--text3);margin-bottom:8px">Divisão do valor total (some 100%)</div>
      <div class="form-row-2" style="margin-bottom:0">
        <div class="form-row" style="margin-bottom:0">
          <label class="form-label" id="f-shared-mine-label">Minha parte (%)</label>
          <input class="form-input" id="f-shared-mine-pct" type="number" min="0" max="100" value="50" oninput="syncSharedPct('mine')">
        </div>
        <div class="form-row" style="margin-bottom:0">
          <label class="form-label" id="f-shared-partner-label">Parte do parceiro (%)</label>
          <input class="form-input" id="f-shared-partner-pct" type="number" min="0" max="100" value="50" oninput="syncSharedPct('partner')">
        </div>
      </div>
      <div style="font-size:11px;color:var(--text3);margin-top:8px" id="f-shared-preview">—</div>
    </div>
    <!-- Parcelamento -->
    <div class="form-row" style="margin-top:4px">
      <label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13px;color:var(--text2)">
        <input type="checkbox" id="f-parceled" onchange="toggleParcelamento()" style="accent-color:var(--green)">
        Pagamento parcelado
      </label>
    </div>
    <div id="f-parcelas-row" style="display:none" class="form-row-2">
      <div class="form-row" style="margin-bottom:0">
        <label class="form-label">Nº de parcelas</label>
        <input class="form-input" id="f-parcelas" type="number" min="2" max="120" value="2" placeholder="Ex: 12">
      </div>
      <div class="form-row" style="margin-bottom:0">
        <label class="form-label">Valor de cada parcela</label>
        <input class="form-input" id="f-parcela-valor" type="number" min="0.01" step="0.01" placeholder="Calculado automaticamente" readonly style="opacity:.7">
      </div>
    </div>
    <input type="hidden" id="f-editing-id">
    <div class="modal-actions"><button class="btn" onclick="closeModal()">Cancelar</button><button class="btn btn-primary" onclick="saveTransaction()">Salvar</button></div>
  </div>
</div>

<script>
// ===================== LOCK SCREEN (login por perfil) =====================
let _pendingProfile=null;
function pickProfile(p){
  _pendingProfile=p;
  document.querySelectorAll('.profile-opt').forEach(el=>el.style.borderColor='var(--border2)');
  document.getElementById('profile-opt-'+p).style.borderColor='var(--green)';
  document.getElementById('lock-sub').textContent='Senha de '+profileName(p);
  document.getElementById('lock-pass-row').style.display='block';
  document.getElementById('lock-input').value='';
  document.getElementById('lock-input').focus();
}
function backToProfilePick(e){
  if(e)e.preventDefault();
  _pendingProfile=null;
  document.getElementById('lock-pass-row').style.display='none';
  document.getElementById('lock-sub').textContent='Quem está entrando?';
  document.getElementById('lock-err').textContent='';
  document.querySelectorAll('.profile-opt').forEach(el=>el.style.borderColor='var(--border2)');
}
function checkPassword(){
  if(!_pendingProfile)return;
  const val=document.getElementById('lock-input').value;
  const inp=document.getElementById('lock-input');
  if(val===PROFILES[_pendingProfile].password){
    _currentProfile=_pendingProfile;
    _viewOwner=_currentProfile; // por padrão, cada um entra vendo a própria área
    localStorage.setItem(DEVICE_PROFILE_KEY,_currentProfile);
    document.getElementById('lock-screen').style.display='none';
    document.getElementById('app-root').style.display='flex';
    document.getElementById('lock-input').value='';
    updateProfileUI();
    refreshAll();
  } else {
    inp.classList.add('err');
    document.getElementById('lock-err').textContent='Senha incorreta. Tente novamente.';
    setTimeout(()=>{inp.classList.remove('err');document.getElementById('lock-err').textContent=''},1800);
    inp.value='';inp.focus();
  }
}
function logoutProfile(){
  _currentProfile=null;_pendingProfile=null;
  document.getElementById('app-root').style.display='none';
  document.getElementById('lock-screen').style.display='flex';
  backToProfilePick();
  // pré-seleciona o último perfil usado neste aparelho, por conveniência
  const last=localStorage.getItem(DEVICE_PROFILE_KEY);
  if(last&&PROFILES[last])pickProfile(last);
}

// ===================== STORAGE =====================
const DB_KEY='fintrack_v2',CATS_KEY='fintrack_cats_v2',IMP_KEY='fintrack_imports',FB_KEY='fintrack_firebase_url';
const store={
  async get(k){if(window.storage){try{const r=await window.storage.get(k);return r?r.value:null}catch{}}return localStorage.getItem(k)},
  async set(k,v){if(window.storage){try{await window.storage.set(k,v);return}catch{}}localStorage.setItem(k,v)}
};

const DEFAULT_CATS={
  expense:[
    {id:'moradia',name:'Moradia',icon:'🏠',color:'#60a5fa',builtin:true},
    {id:'alimentacao',name:'Alimentação',icon:'🍽️',color:'#fb923c',builtin:true},
    {id:'transporte',name:'Transporte',icon:'🚗',color:'#a78bfa',builtin:true},
    {id:'saude',name:'Saúde',icon:'❤️',color:'#f87171',builtin:true},
    {id:'educacao',name:'Educação',icon:'📚',color:'#34d399',builtin:true},
    {id:'lazer',name:'Lazer',icon:'🎮',color:'#fbbf24',builtin:true},
    {id:'vestuario',name:'Vestuário',icon:'👕',color:'#c084fc',builtin:true},
    {id:'servicos',name:'Serviços',icon:'📱',color:'#38bdf8',builtin:true},
    {id:'outros_d',name:'Outros',icon:'📦',color:'#94a3b8',builtin:true},
  ],
  income:[
    {id:'salario',name:'Salário',icon:'💼',color:'#22d3a0',builtin:true},
    {id:'freelance',name:'Freelance',icon:'💻',color:'#60a5fa',builtin:true},
    {id:'investimento',name:'Investimento',icon:'📈',color:'#fbbf24',builtin:true},
    {id:'bonus',name:'Bônus',icon:'🎁',color:'#f472b6',builtin:true},
    {id:'outros_r',name:'Outros',icon:'💰',color:'#94a3b8',builtin:true},
  ]
};

// ===================== PERFIS (casal) =====================
const PROFILES={
  thiago:{id:'thiago',name:'Thiago',icon:'👤',password:'contths123'},
  luiza:{id:'luiza',name:'Luiza',icon:'👤',password:'contlcr123'}
};
const DEVICE_PROFILE_KEY='fintrack_device_profile'; // só local, não sincroniza
let _currentProfile=null; // 'thiago' | 'luiza'
let _viewOwner='casal'; // 'thiago' | 'luiza' | 'casal' — o que está sendo exibido nas telas

function otherProfile(p){return p==='thiago'?'luiza':'thiago'}
function profileName(p){return PROFILES[p]?PROFILES[p].name:p}

function updateProfileUI(){
  const badge=document.getElementById('profile-badge');
  if(badge)badge.textContent=(PROFILES[_currentProfile]?.icon||'')+' '+profileName(_currentProfile);
  ['thiago','luiza','casal'].forEach(v=>{
    const el=document.getElementById('view-'+v);
    if(!el)return;
    const active=_viewOwner===v;
    el.style.background=active?'var(--green)':'var(--bg3)';
    el.style.color=active?'#0f1117':'var(--text2)';
    el.style.borderColor=active?'var(--green)':'var(--border2)';
    el.style.fontWeight=active?'600':'400';
  });
}
function viewSuffix(){
  if(_viewOwner==='casal')return ' · Casal';
  return ' · '+profileName(_viewOwner);
}
function setViewOwner(v){
  _viewOwner=v;
  updateProfileUI();
  refreshAll();
}

// retorna as transações visíveis de acordo com a visão selecionada (Meu/Dela/Casal)
function getVisibleTx(){
  const all=_db.transactions;
  if(_viewOwner==='casal')return all;
  return all.filter(t=>t.owner===_viewOwner);
}
// retorna um "db" com transactions já filtradas pela visão — usar em telas de leitura/relatório
function visibleDB(){return{transactions:getVisibleTx()}}

let _db={transactions:[]},_cats=JSON.parse(JSON.stringify(DEFAULT_CATS)),_imports=[],_fbUrl='';

async function loadAll(){
  try{
    const[dbR,catsR,impsR,fbR]=await Promise.all([store.get(DB_KEY),store.get(CATS_KEY),store.get(IMP_KEY),store.get(FB_KEY)]);
    _db=dbR?JSON.parse(dbR):{transactions:[]};
    _cats=catsR?JSON.parse(catsR):JSON.parse(JSON.stringify(DEFAULT_CATS));
    _imports=impsR?JSON.parse(impsR):[];
    _fbUrl=fbR||'';
    await migrateDoacaoToExpense();
    await migrateOwnerField();
  }catch{}
}

// Migração: dados criados antes da existência de perfis (Thiago/Luiza) não têm "owner".
// Como a base já existente pertence ao Thiago (só ele usava o app), atribuímos owner='thiago'
// a tudo que ainda não tiver dono definido. Roda uma única vez, silenciosamente.
async function migrateOwnerField(){
  let changed=false;
  _db.transactions.forEach(t=>{
    if(!t.owner){t.owner='thiago';changed=true}
  });
  if(changed)await saveDB(_db);
}

// Migração: categorias "Doação" criadas anteriormente como receita passam a ser despesa
async function migrateDoacaoToExpense(){
  const norm=s=>String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  let changed=false;
  const doacaoIncomeCats=_cats.income.filter(c=>/doa[cç][aã]o/.test(norm(c.name)));
  if(doacaoIncomeCats.length){
    doacaoIncomeCats.forEach(c=>{
      _cats.income=_cats.income.filter(x=>x.id!==c.id);
      if(!_cats.expense.find(x=>x.id===c.id)) _cats.expense.push(c);
      changed=true;
    });
  }
  const doacaoIds=new Set(_cats.expense.filter(c=>/doa[cç][aã]o/.test(norm(c.name))).map(c=>c.id));
  if(doacaoIds.size){
    _db.transactions.forEach(t=>{
      if(doacaoIds.has(t.category)&&t.type==='income'){t.type='expense';changed=true}
    });
  }
  if(changed){await saveCats(_cats);await saveDB(_db)}
}
function loadCats(){return _cats}
function loadDB(){return _db}
function loadImports(){return _imports}
async function saveCats(c){_cats=c;await store.set(CATS_KEY,JSON.stringify(c))}
async function saveDB(db){_db=db;await store.set(DB_KEY,JSON.stringify(db))}
async function saveImports(a){_imports=a;await store.set(IMP_KEY,JSON.stringify(a))}

// auto-sync to cloud on every save
async function saveDBAndSync(db){
  await saveDB(db);
  if(_fbUrl) autoSyncToCloud();
}
async function autoSyncToCloud(){
  try{
    const payload={db:JSON.stringify(_db),cats:JSON.stringify(_cats),imports:JSON.stringify(_imports),updated:new Date().toISOString()};
    await fetch(_fbUrl+'/fintrack.json',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    logCloud('☁ Auto-sync: '+new Date().toLocaleTimeString('pt-BR'));
  }catch(e){logCloud('⚠ Auto-sync falhou: '+e.message)}
}

function allCats(){const c=loadCats();return[...c.expense,...c.income]}
function getCat(id){return allCats().find(c=>c.id===id)||{name:id||'—',icon:'📦',color:'#888'}}
function slugify(s){return s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,'').substring(0,30)}
function randomColor(){const p=['#f87171','#fb923c','#fbbf24','#34d399','#60a5fa','#a78bfa','#f472b6','#38bdf8','#22d3a0','#c084fc'];return p[Math.floor(Math.random()*p.length)]}

const MONTHS=['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const MONTHS_SHORT=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
function fmt(v){return 'R$ '+Math.abs(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}

function getYears(){
  const db=loadDB(),now=new Date().getFullYear(),ys=new Set([now]);
  db.transactions.forEach(t=>ys.add(new Date(t.date+'T12:00').getFullYear()));
  return[...ys].sort((a,b)=>b-a);
}
function populateMonthSelect(id,cur){
  const sel=document.getElementById(id);if(!sel)return;
  sel.innerHTML=(id.startsWith('tx-')||id==='exp-month'?'<option value="-1">Todos os meses</option>':'')+MONTHS.map((m,i)=>`<option value="${i}"${i===cur?' selected':''}>${m}</option>`).join('');
}
function populateYearSelect(id,cur){
  const sel=document.getElementById(id);if(!sel)return;
  // preserva o ano atualmente selecionado, se existir
  const selected=sel.value?parseInt(sel.value):cur;
  const years=getYears();
  // se o ano selecionado não existe na lista nova, usa o preferido (cur)
  const chosen=years.includes(selected)?selected:cur;
  sel.innerHTML=years.map(y=>`<option value="${y}"${y===chosen?' selected':''}>${y}</option>`).join('');
}
function filterTx(txs,month,year){
  return txs.filter(t=>{const d=new Date(t.date+'T12:00');return d.getFullYear()===year&&(month===-1||d.getMonth()===month)});
}

// ===================== CHARTS =====================
let chartMonthly=null,chartPie=null,chartCatPie=null,chartAnual=null,chartCatAnual=null,chartEvolucao=null,chartAcumulado=null,chartInsCat=null;
function dc(c){if(c){try{c.destroy()}catch{}}return null}

// ===================== SECTIONS =====================
function showSection(name){
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  const sec=document.getElementById('sec-'+name);
  if(sec)sec.classList.add('active');
  const ni=document.querySelector(`.nav-item[data-section="${name}"]`);
  if(ni)ni.classList.add('active');
  document.querySelectorAll('.mob-btn').forEach(b=>b.classList.remove('active'));
  const mm={dashboard:'mob-dashboard',transacoes:'mob-transacoes',evolucao:'mob-evolucao'};
  if(mm[name])document.getElementById(mm[name])?.classList.add('active');
  document.getElementById('sidebar').classList.remove('open');
  if(name==='dashboard')updateDashboard();
  if(name==='transacoes')renderTxList();
  if(name==='evolucao')renderEvolucao();
  if(name==='comparativo')renderComparativo();
  if(name==='insights')renderInsights();
  if(name==='patrimonio')renderPatrimonio();
  if(name==='categorias')renderCategorias();
  if(name==='anual')renderAnual();
  if(name==='importar')renderImportHistory();
  if(name==='exportar')renderExportSection();
  if(name==='gcat')renderCatManager();
  if(name==='cloud')renderCloudSection();
  if(name==='reset')renderResetSection();
}
function toggleMobileMenu(){document.getElementById('sidebar').classList.toggle('open')}

// ===================== MODAL =====================
let currentType='expense';
function openModal(id){
  document.getElementById('modal-bg').classList.add('open');
  document.getElementById('modal-title').textContent=id?'Editar transação':'Nova transação';
  document.getElementById('f-editing-id').value=id||'';
  // reset parcelamento
  document.getElementById('f-parceled').checked=false;
  document.getElementById('f-parcelas-row').style.display='none';
  document.getElementById('f-parcelas').value=2;
  document.getElementById('f-parcela-valor').value='';
  // hide parcelamento when editing
  document.getElementById('f-parceled').closest('.form-row').style.display=id?'none':'block';
  document.getElementById('f-parcelas-row').style.display='none';
  // reset compartilhado
  document.getElementById('f-shared').checked=false;
  document.getElementById('f-shared-split-row').style.display='none';
  document.getElementById('f-shared-mine-pct').value=50;
  document.getElementById('f-shared-partner-pct').value=50;
  document.getElementById('f-shared-partner-name').textContent=profileName(otherProfile(_currentProfile));
  document.getElementById('f-shared-mine-label').textContent='Minha parte ('+profileName(_currentProfile)+') %';
  document.getElementById('f-shared-partner-label').textContent='Parte de '+profileName(otherProfile(_currentProfile))+' (%)';
  if(id){
    const t=loadDB().transactions.find(x=>x.id===id);if(!t)return;
    setType(t.type);
    document.getElementById('f-desc').value=t.description;
    document.getElementById('f-amount').value=t.amount;
    document.getElementById('f-date').value=t.date;
    const c=getCat(t.category);
    document.getElementById('f-cat').value=t.category;
    document.getElementById('f-cat-search').value=c.icon+' '+c.name;
    // não permite alterar o compartilhamento na edição — só na criação
    document.getElementById('f-shared-row').style.display='none';
    if(t.shared){
      document.getElementById('f-shared-split-row').style.display='block';
      document.getElementById('f-shared-split-row').innerHTML=`<div style="font-size:12px;color:var(--text2)">🤝 Esta transação faz parte de um gasto compartilhado (total ${fmt(t.sharedTotal||0)}, pago por ${profileName(t.paidBy)}). Editar aqui só ajusta esta linha — a parte do(a) ${profileName(otherProfile(t.owner))} não muda automaticamente. Se precisar desfazer a divisão, exclua e lance novamente.</div>`;
    }
  }else{
    document.getElementById('f-date').value=new Date().toISOString().split('T')[0];
    document.getElementById('f-desc').value='';
    document.getElementById('f-amount').value='';
    document.getElementById('f-shared-row').style.display='block';
    setType('expense');
  }
  // auto-calc parcela when amount changes
  document.getElementById('f-amount').oninput=()=>{if(document.getElementById('f-parceled').checked)calcParcelaValor();updateSharedPreview()};
  document.getElementById('f-parcelas').oninput=calcParcelaValor;
  updateSharedPreview();
}
function closeModal(){document.getElementById('modal-bg').classList.remove('open')}
function closeModalBg(e){if(e.target.id==='modal-bg')closeModal()}

function setType(type){
  currentType=type;
  document.getElementById('rb-income').className='radio-btn'+(type==='income'?' sel-income':'');
  document.getElementById('rb-expense').className='radio-btn'+(type==='expense'?' sel-expense':'');
  document.getElementById('f-cat').value='';
  document.getElementById('f-cat-search').value='';
  renderCatDropdown('');
  // mostra dica de valor negativo apenas para receita
  document.getElementById('f-amount-hint').style.display=type==='income'?'block':'none';
  // parcelamento só faz sentido para despesas
  const pRow=document.getElementById('f-parceled').closest('.form-row');
  if(type==='income'){
    document.getElementById('f-parceled').checked=false;
    document.getElementById('f-parcelas-row').style.display='none';
    pRow.style.display='none';
  } else if(!document.getElementById('f-editing-id').value){
    pRow.style.display='block';
  }
}
function getAllCatsForType(type){const c=loadCats();return[...c[type].map(x=>({...x,_primary:true})),...c[type==='expense'?'income':'expense'].map(x=>({...x,_primary:false}))]}
function renderCatDropdown(filter){
  const dd=document.getElementById('f-cat-dropdown');if(!dd)return;
  const all=getAllCatsForType(currentType);
  const q=filter.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  const filtered=all.filter(c=>c.name.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').includes(q));
  if(!filtered.length){dd.innerHTML='<div style="padding:10px 14px;font-size:12px;color:var(--text3)">Nenhuma encontrada</div>';return}
  const pri=filtered.filter(c=>c._primary),sec=filtered.filter(c=>!c._primary);
  let html=pri.map(c=>catOpt(c)).join('');
  if(sec.length){if(pri.length)html+=`<div style="padding:4px 14px;font-size:10px;color:var(--text3);background:var(--bg4);border-top:1px solid var(--border)">Outras</div>`;html+=sec.map(c=>catOpt(c)).join('')}
  dd.innerHTML=html;
}
function catOpt(c){
  const sel=document.getElementById('f-cat')?.value===c.id;
  return`<div onclick="selectCat('${c.id}','${c.name.replace(/'/g,"\\'")}','${c.icon}')" style="display:flex;align-items:center;gap:9px;padding:9px 14px;cursor:pointer;font-size:13px;background:${sel?'var(--bg4)':'transparent'}" onmouseover="this.style.background='var(--bg4)'" onmouseout="this.style.background='${sel?'var(--bg4)':'transparent'}'"><span style="width:8px;height:8px;border-radius:50%;background:${c.color};flex-shrink:0"></span><span>${c.icon}</span><span>${c.name}</span>${sel?'<span style="margin-left:auto;color:var(--green);font-size:11px">✓</span>':''}</div>`;
}
function selectCat(id,name,icon){document.getElementById('f-cat').value=id;document.getElementById('f-cat-search').value=icon+' '+name;document.getElementById('f-cat-dropdown').style.display='none'}
function showCatDropdown(){document.getElementById('f-cat-dropdown').style.display='block';renderCatDropdown(document.getElementById('f-cat-search').value);setTimeout(()=>document.addEventListener('click',function h(e){if(!document.getElementById('f-cat-dropdown')?.contains(e.target)&&e.target!==document.getElementById('f-cat-search')){document.getElementById('f-cat-dropdown').style.display='none';document.removeEventListener('click',h)}},{once:true}),10)}
function filterCatOptions(){document.getElementById('f-cat-dropdown').style.display='block';renderCatDropdown(document.getElementById('f-cat-search').value)}

function toggleParcelamento(){
  const checked=document.getElementById('f-parceled').checked;
  document.getElementById('f-parcelas-row').style.display=checked?'grid':'none';
  if(checked) calcParcelaValor();
}
function calcParcelaValor(){
  const total=parseFloat(document.getElementById('f-amount').value)||0;
  const n=parseInt(document.getElementById('f-parcelas').value)||2;
  document.getElementById('f-parcela-valor').value=n>0?(total/n).toFixed(2):'';
}

function toggleShared(){
  const checked=document.getElementById('f-shared').checked;
  document.getElementById('f-shared-split-row').style.display=checked?'block':'none';
  updateSharedPreview();
}
function syncSharedPct(which){
  const mineEl=document.getElementById('f-shared-mine-pct'),partnerEl=document.getElementById('f-shared-partner-pct');
  let mine=parseFloat(mineEl.value)||0,partner=parseFloat(partnerEl.value)||0;
  if(which==='mine')partner=100-mine; else mine=100-partner;
  mine=Math.max(0,Math.min(100,mine));partner=Math.max(0,Math.min(100,partner));
  mineEl.value=mine;partnerEl.value=partner;
  updateSharedPreview();
}
function updateSharedPreview(){
  const prev=document.getElementById('f-shared-preview');if(!prev)return;
  if(!document.getElementById('f-shared').checked){prev.textContent='—';return}
  const total=parseFloat(document.getElementById('f-amount').value)||0;
  const mine=parseFloat(document.getElementById('f-shared-mine-pct').value)||0;
  const partner=parseFloat(document.getElementById('f-shared-partner-pct').value)||0;
  prev.textContent=`Total ${fmt(total)} → ${profileName(_currentProfile)}: ${fmt(total*mine/100)} (${mine}%) · ${profileName(otherProfile(_currentProfile))}: ${fmt(total*partner/100)} (${partner}%)`;
}

async function saveTransaction(){
  const desc=document.getElementById('f-desc').value.trim();
  const amount=parseFloat(document.getElementById('f-amount').value);
  const date=document.getElementById('f-date').value;
  const cat=document.getElementById('f-cat').value;
  if(!desc||!date||isNaN(amount)||amount===0){alert('Preencha todos os campos.');return}
  if(currentType==='expense'&&amount<0){alert('Despesas devem ter valor positivo.');return}
  if(!cat){alert('Selecione uma categoria.');return}
  const db=loadDB();
  const editingId=document.getElementById('f-editing-id').value;
  const isParceled=document.getElementById('f-parceled').checked&&!editingId;
  const isShared=document.getElementById('f-shared').checked&&!editingId&&!isParceled;
  if(isShared){
    const minePct=parseFloat(document.getElementById('f-shared-mine-pct').value)||0;
    const partnerPct=parseFloat(document.getElementById('f-shared-partner-pct').value)||0;
    if(Math.round(minePct+partnerPct)!==100){alert('A divisão precisa somar 100%.');return}
    const partner=otherProfile(_currentProfile);
    const groupId='sh_'+Date.now().toString(36);
    db.transactions.unshift({
      id:Date.now().toString(),type:currentType,description:desc,
      amount:parseFloat((amount*minePct/100).toFixed(2)),date,category:cat,
      owner:_currentProfile,shared:true,groupId,sharedTotal:amount,splitPct:minePct,paidBy:_currentProfile
    });
    db.transactions.unshift({
      id:Date.now().toString()+'_p',type:currentType,description:desc,
      amount:parseFloat((amount*partnerPct/100).toFixed(2)),date,category:cat,
      owner:partner,shared:true,groupId,sharedTotal:amount,splitPct:partnerPct,paidBy:_currentProfile
    });
  } else if(isParceled){
    const n=parseInt(document.getElementById('f-parcelas').value)||2;
    if(n<2||n>120){alert('Número de parcelas deve ser entre 2 e 120.');return}
    const parcelaValor=amount/n;
    const baseDate=new Date(date+'T12:00');
    const groupId='parc_'+Date.now().toString(36);
    for(let i=0;i<n;i++){
      const d=new Date(baseDate);
      d.setMonth(d.getMonth()+i);
      const ds=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
      db.transactions.unshift({
        id:Date.now().toString()+Math.random().toString(36).slice(2),
        type:currentType,
        description:`${desc} (${i+1}/${n})`,
        amount:parseFloat(parcelaValor.toFixed(2)),
        date:ds,
        category:cat,
        parcelaGroup:groupId,
        parcelaNum:i+1,
        parcelaTotal:n,
        owner:_currentProfile
      });
    }
  } else {
    if(editingId){
      const idx=db.transactions.findIndex(t=>t.id===editingId);
      if(idx>=0)db.transactions[idx]={...db.transactions[idx],type:currentType,description:desc,amount,date,category:cat};
    } else {
      db.transactions.unshift({id:Date.now().toString(),type:currentType,description:desc,amount,date,category:cat,owner:_currentProfile});
    }
  }
  db.transactions.sort((a,b)=>new Date(b.date)-new Date(a.date));
  await saveDBAndSync(db);
  closeModal();populateAllSelects();refreshAll();
}
async function delTx(id){
  const db=loadDB();
  const t=db.transactions.find(x=>x.id===id);
  if(!t)return;
  let idsToDelete=[id];
  if(t.shared&&t.groupId){
    const alsoDeletePartner=confirm('Esta transação é compartilhada. Deseja excluir também a parte de '+profileName(otherProfile(t.owner))+'?\n\nOK = excluir os dois lados\nCancelar = excluir só a sua parte');
    if(alsoDeletePartner){
      idsToDelete=db.transactions.filter(x=>x.groupId===t.groupId).map(x=>x.id);
    }
  } else {
    if(!confirm('Excluir esta transação?'))return;
  }
  db.transactions=db.transactions.filter(x=>!idsToDelete.includes(x.id));
  await saveDBAndSync(db);populateAllSelects();refreshAll();
}

// ===================== DASHBOARD =====================
function updateDashboard(){
  const now=new Date();
  const month=parseInt(document.getElementById('dash-month')?.value??now.getMonth());
  const year=parseInt(document.getElementById('dash-year')?.value??now.getFullYear());
  document.getElementById('dash-period-label').textContent=MONTHS[month]+' de '+year+viewSuffix();
  const db=visibleDB(),txs=filterTx(db.transactions,month,year);
  const income=txs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);
  const expense=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
  const balance=income-expense;
  document.getElementById('c-income').textContent=fmt(income);
  document.getElementById('c-income-count').textContent=txs.filter(t=>t.type==='income').length+' entrada(s)';
  document.getElementById('c-expense').textContent=fmt(expense);
  document.getElementById('c-expense-count').textContent=txs.filter(t=>t.type==='expense').length+' saída(s)';
  document.getElementById('c-balance').textContent=fmt(balance);
  document.getElementById('c-balance').style.color=balance>=0?'var(--green)':'var(--red)';
  document.getElementById('c-balance-pct').textContent=income>0?'Poupança: '+Math.round((balance/income)*100)+'%':'—';
  const catMap={};txs.filter(t=>t.type==='expense').forEach(t=>{catMap[t.category]=(catMap[t.category]||0)+t.amount});
  const top=Object.entries(catMap).sort((a,b)=>b[1]-a[1])[0];
  if(top){const c=getCat(top[0]);document.getElementById('c-topcat').textContent=c.icon+' '+c.name;document.getElementById('c-topcat-amt').textContent=fmt(top[1])}
  else{document.getElementById('c-topcat').textContent='—';document.getElementById('c-topcat-amt').textContent='Sem despesas'}
  const last6=[];
  for(let i=5;i>=0;i--){let m=month-i,y=year;if(m<0){m+=12;y--}const t=filterTx(db.transactions,m,y);last6.push({label:MONTHS_SHORT[m],income:t.filter(x=>x.type==='income').reduce((s,x)=>s+x.amount,0),expense:t.filter(x=>x.type==='expense').reduce((s,x)=>s+x.amount,0)})}
  chartMonthly=dc(chartMonthly);
  chartMonthly=new Chart(document.getElementById('chart-monthly'),{type:'bar',data:{labels:last6.map(x=>x.label),datasets:[{label:'Receitas',data:last6.map(x=>x.income),backgroundColor:'rgba(34,211,160,0.7)',borderRadius:4},{label:'Despesas',data:last6.map(x=>x.expense),backgroundColor:'rgba(245,101,101,0.7)',borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}});
  const expCats=allCats().filter(c=>loadCats().expense.find(e=>e.id===c.id)).map(c=>({cat:c,total:txs.filter(t=>t.type==='expense'&&t.category===c.id).reduce((s,t)=>s+t.amount,0)})).filter(x=>x.total>0).sort((a,b)=>b.total-a.total);
  chartPie=dc(chartPie);
  const pieLeg=document.getElementById('pie-legend');
  if(!expCats.length){pieLeg.innerHTML='<div style="color:var(--text3);font-size:12px">Sem despesas</div>';return}
  chartPie=new Chart(document.getElementById('chart-pie'),{type:'doughnut',data:{labels:expCats.map(x=>x.cat.name),datasets:[{data:expCats.map(x=>x.total),backgroundColor:expCats.map(x=>x.cat.color),borderWidth:0,hoverOffset:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>' '+fmt(ctx.raw)}}}}});
  pieLeg.innerHTML=expCats.map(x=>`<span class="legend-item"><span class="legend-dot" style="background:${x.cat.color}"></span><span style="color:var(--text2)">${x.cat.name}</span><span style="margin-left:auto;font-family:DM Mono;font-size:12px">${fmt(x.total)}</span></span>`).join('');
  pieLeg.style.cssText='display:flex;flex-direction:column;gap:7px;margin-top:14px';
  renderTxItems(document.getElementById('recent-list'),db.transactions.slice(0,8));
  renderBalanceCard(month,year);
}

// ===================== BALANÇO DO CASAL =====================
// Considera SEMPRE os dados completos (não é afetado pelo filtro Meu/Dela/Casal),
// pois o balanço é, por natureza, uma visão conjunta.
function renderBalanceCard(month,year){
  const el=document.getElementById('balance-card-body');
  if(!el)return;
  const monthTx=filterTx(_db.transactions,month,year).filter(t=>t.shared&&t.groupId);
  const seen=new Set();
  let owedToThiago=0, owedToLuiza=0, totalShared=0;
  monthTx.forEach(t=>{
    if(seen.has(t.groupId))return;
    seen.add(t.groupId);
    const group=_db.transactions.filter(x=>x.groupId===t.groupId);
    const payer=t.paidBy;
    if(!payer||group.length<2)return;
    totalShared+=t.sharedTotal||0;
    const otherLine=group.find(x=>x.owner!==payer);
    if(!otherLine)return;
    if(payer==='thiago')owedToThiago+=otherLine.amount;
    else if(payer==='luiza')owedToLuiza+=otherLine.amount;
  });
  if(!seen.size){
    el.innerHTML='<div class="empty" style="padding:14px 0"><div class="empty-icon">🤝</div>Nenhum gasto compartilhado neste mês ainda.</div>';
    return;
  }
  const net=owedToThiago-owedToLuiza; // >0: Luiza deve a Thiago; <0: Thiago deve a Luiza
  let resultLine;
  if(Math.abs(net)<0.01){
    resultLine=`<span style="color:var(--green);font-weight:600">✅ Contas equilibradas — ninguém deve nada este mês.</span>`;
  } else if(net>0){
    resultLine=`<span style="color:var(--amber);font-weight:600">Luiza deve ${fmt(net)} a Thiago</span>`;
  } else {
    resultLine=`<span style="color:var(--amber);font-weight:600">Thiago deve ${fmt(Math.abs(net))} a Luiza</span>`;
  }
  el.innerHTML=`
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:14px">
      <div><div style="font-size:11px;color:var(--text3);margin-bottom:3px">Total pago por Thiago (compartilhado)</div><div style="font-family:DM Mono;font-size:15px;font-weight:600">${fmt((_db.transactions.filter(t=>t.shared&&t.paidBy==='thiago'&&filterTx([t],month,year).length).reduce((s,t2)=>s+ (t2.sharedTotal&&t2.owner===t2.paidBy?t2.sharedTotal:0),0)))}</div></div>
      <div><div style="font-size:11px;color:var(--text3);margin-bottom:3px">Total pago por Luiza (compartilhado)</div><div style="font-family:DM Mono;font-size:15px;font-weight:600">${fmt((_db.transactions.filter(t=>t.shared&&t.paidBy==='luiza'&&filterTx([t],month,year).length).reduce((s,t2)=>s+ (t2.sharedTotal&&t2.owner===t2.paidBy?t2.sharedTotal:0),0)))}</div></div>
    </div>
    <div style="padding:12px 14px;background:var(--bg3);border-radius:var(--radius-sm);font-size:14px">${resultLine}</div>
  `;
}

// ===================== TX LIST com filtros =====================
let txTypeFilter='all', txCatFilters=new Set();

function setTxTypeFilter(type){
  txTypeFilter=type;
  ['all','income','expense'].forEach(t=>{
    const chip=document.getElementById('chip-type-'+t);
    if(!chip)return;
    chip.className='filter-chip'+(type===t?(t==='all'?' active':t==='income'?' active-income':' active-expense'):'');
  });
  renderTxList();
}

function toggleCatFilter(catId){
  if(txCatFilters.has(catId))txCatFilters.delete(catId);
  else txCatFilters.add(catId);
  renderTxList();
}

function buildCatFilterChips(){
  const listOwner=_viewOwner==='casal'?_currentProfile:_viewOwner;
  const usedCats=[...new Set(_db.transactions.filter(t=>t.owner===listOwner).map(t=>t.category))];
  const container=document.getElementById('cat-filter-chips');
  if(!container)return;
  container.innerHTML=usedCats.map(cid=>{
    const c=getCat(cid),active=txCatFilters.has(cid);
    return`<div class="filter-chip ${active?'active':''}" onclick="toggleCatFilter('${cid}')" id="catfc-${cid}"><span style="width:7px;height:7px;border-radius:50%;background:${c.color};display:inline-block"></span>${c.icon} ${c.name}${active?'<span class="chip-x">✕</span>':''}</div>`;
  }).join('');
}

function renderTxList(){
  buildCatFilterChips();
  const month=parseInt(document.getElementById('tx-month')?.value??-1);
  const year=parseInt(document.getElementById('tx-year')?.value??new Date().getFullYear());
  // A lista de transações nunca é "do casal" — se a visão global estiver em Casal,
  // mostramos a lista da pessoa atualmente logada (cada gasto só existe dentro de uma pessoa).
  const listOwner=_viewOwner==='casal'?_currentProfile:_viewOwner;
  const db={transactions:_db.transactions.filter(t=>t.owner===listOwner)};
  let txs=month===-1?db.transactions.filter(t=>new Date(t.date+'T12:00').getFullYear()===year):filterTx(db.transactions,month,year);
  if(txTypeFilter!=='all')txs=txs.filter(t=>t.type===txTypeFilter);
  if(txCatFilters.size>0)txs=txs.filter(t=>txCatFilters.has(t.category));
  const lbl=document.getElementById('tx-count-label');
  if(lbl)lbl.textContent=txs.length+' transação(ões) · '+profileName(listOwner)+(_viewOwner==='casal'?' (visão Casal não combina listas — mostrando as suas)':'');
  const title=document.getElementById('tx-list-title');
  if(title)title.textContent=txs.length+' transação(ões) encontrada(s)';
  // mostrar botão excluir filtradas só quando há filtro ativo
  const hasFilter=txTypeFilter!=='all'||txCatFilters.size>0||month!==-1;
  const delBtn=document.getElementById('btn-delete-filtered');
  if(delBtn)delBtn.style.display=(hasFilter&&txs.length>0)?'inline-flex':'none';
  renderTxItems(document.getElementById('tx-list'),txs);
}

async function deleteFiltered(){
  const month=parseInt(document.getElementById('tx-month')?.value??-1);
  const year=parseInt(document.getElementById('tx-year')?.value??new Date().getFullYear());
  const listOwner=_viewOwner==='casal'?_currentProfile:_viewOwner;
  const db={transactions:_db.transactions.filter(t=>t.owner===listOwner)};
  let toDelete=month===-1?db.transactions.filter(t=>new Date(t.date+'T12:00').getFullYear()===year):filterTx(db.transactions,month,year);
  if(txTypeFilter!=='all')toDelete=toDelete.filter(t=>t.type===txTypeFilter);
  if(txCatFilters.size>0)toDelete=toDelete.filter(t=>txCatFilters.has(t.category));
  if(!toDelete.length){alert('Nenhuma transação selecionada.');return}
  if(!confirm(`Excluir ${toDelete.length} transação(ões) filtrada(s)? Esta ação não pode ser desfeita.`))return;
  const deleteIds=new Set(toDelete.map(t=>t.id));
  const fullDb=loadDB();fullDb.transactions=fullDb.transactions.filter(t=>!deleteIds.has(t.id));
  await saveDBAndSync(fullDb);
  refreshAll();
}

function renderTxItems(container,txs){
  if(!container)return;
  if(!txs.length){container.innerHTML='<div class="empty"><div class="empty-icon">🔍</div>Nenhuma transação encontrada</div>';return}
  container.innerHTML=txs.map(t=>{
    const c=getCat(t.category),d=new Date(t.date+'T12:00');
    const isNegIncome=t.type==='income'&&t.amount<0;
    const amtColor=t.type==='expense'?'var(--red)':(isNegIncome?'var(--amber)':'var(--green)');
    const sign=t.amount<0?'-':(t.type==='income'?'+':'-');
    const ownerTag=t.owner?`<span style="font-size:10px;padding:1px 6px;border-radius:10px;background:var(--bg4);color:var(--text3);margin-left:6px">${profileName(t.owner)}</span>`:'';
    const sharedTag=t.shared?`<span style="font-size:10px;padding:1px 6px;border-radius:10px;background:var(--green-bg);color:var(--green);margin-left:4px" title="Total ${fmt(t.sharedTotal||0)} · pago por ${profileName(t.paidBy)} · ${t.splitPct}%">🤝 ${t.splitPct}%</span>`:'';
    return`<div class="tx-item"><div class="tx-icon" style="background:${c.color}22">${c.icon}</div><div class="tx-info"><div class="tx-name">${t.description}${ownerTag}${sharedTag}</div><div class="tx-cat">${c.name}${isNegIncome?' · abatimento':''}</div></div><div class="tx-date">${d.toLocaleDateString('pt-BR')}</div><div class="tx-amount" style="color:${amtColor}">${sign} ${fmt(Math.abs(t.amount))}</div><div class="tx-actions"><button class="tx-btn" onclick="openModal('${t.id}')" title="Editar">✏️</button><button class="tx-btn" onclick="delTx('${t.id}')" title="Excluir">🗑</button></div></div>`;
  }).join('');
}

// ===================== EVOLUÇÃO =====================
function renderEvolucao(){
  const year=parseInt(document.getElementById('ev-year')?.value??new Date().getFullYear());
  const db=visibleDB();
  const months=MONTHS.map((name,m)=>{
    const txs=filterTx(db.transactions,m,year);
    const income=txs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);
    const expense=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
    return{name,short:MONTHS_SHORT[m],income,expense,balance:income-expense,count:txs.length};
  });
  // acumulado
  let acc=0;const accumulated=months.map(m=>{acc+=m.balance;return acc});
  // cards
  const bestIncome=months.reduce((a,b)=>b.income>a.income?b:a,months[0]);
  const worstExpense=months.reduce((a,b)=>b.expense>a.expense?b:a,months[0]);
  document.getElementById('ev-best-income').textContent=bestIncome.income>0?bestIncome.name:'—';
  document.getElementById('ev-best-income-val').textContent=bestIncome.income>0?fmt(bestIncome.income):'Sem receitas';
  document.getElementById('ev-worst-expense').textContent=worstExpense.expense>0?worstExpense.name:'—';
  document.getElementById('ev-worst-expense-val').textContent=worstExpense.expense>0?fmt(worstExpense.expense):'Sem despesas';
  document.getElementById('ev-accumulated').textContent=fmt(acc);
  document.getElementById('ev-accumulated').style.color=acc>=0?'var(--green)':'var(--red)';
  document.getElementById('ev-accumulated-sub').textContent='Saldo acumulado em '+year;
  // médias
  const activeMths=months.filter(m=>m.count>0);
  const avgInc=activeMths.length?activeMths.reduce((s,m)=>s+m.income,0)/activeMths.length:0;
  const avgExp=activeMths.length?activeMths.reduce((s,m)=>s+m.expense,0)/activeMths.length:0;
  const avgBal=activeMths.length?activeMths.reduce((s,m)=>s+m.balance,0)/activeMths.length:0;
  document.getElementById('ev-medias').innerHTML=[
    {label:'Receita média mensal',value:fmt(avgInc),color:'var(--green)'},
    {label:'Despesa média mensal',value:fmt(avgExp),color:'var(--red)'},
    {label:'Saldo médio mensal',value:fmt(avgBal),color:avgBal>=0?'var(--green)':'var(--red)'},
    {label:'Meses com movimentação',value:activeMths.length+' de 12',color:'var(--text)'},
  ].map(r=>`<div style="display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid var(--border)"><span style="font-size:13px;color:var(--text2)">${r.label}</span><span style="font-family:DM Mono;font-size:13px;color:${r.color}">${r.value}</span></div>`).join('');
  // gráfico principal
  chartEvolucao=dc(chartEvolucao);
  chartEvolucao=new Chart(document.getElementById('chart-evolucao'),{
    type:'bar',
    data:{labels:MONTHS_SHORT,datasets:[
      {label:'Receitas',data:months.map(m=>m.income),backgroundColor:'rgba(34,211,160,0.65)',borderRadius:4},
      {label:'Despesas',data:months.map(m=>m.expense),backgroundColor:'rgba(245,101,101,0.65)',borderRadius:4},
      {type:'line',label:'Saldo',data:months.map(m=>m.balance),borderColor:'rgba(96,165,250,0.9)',backgroundColor:'transparent',borderWidth:2,pointRadius:3,tension:0.3},
    ]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}
  });
  // gráfico acumulado
  chartAcumulado=dc(chartAcumulado);
  chartAcumulado=new Chart(document.getElementById('chart-acumulado'),{
    type:'line',
    data:{labels:MONTHS_SHORT,datasets:[{label:'Saldo acumulado',data:accumulated,borderColor:'rgba(251,191,36,0.9)',backgroundColor:'rgba(251,191,36,0.08)',borderWidth:2,pointRadius:4,tension:0.3,fill:true}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}
  });
}

// ===================== INSIGHTS =====================
function renderInsights(){
  const year=parseInt(document.getElementById('ins-year')?.value??new Date().getFullYear());
  const db=visibleDB();
  const allTxs=db.transactions.filter(t=>new Date(t.date+'T12:00').getFullYear()===year);
  // year totals
  const months=MONTHS.map((name,m)=>{
    const txs=filterTx(db.transactions,m,year);
    const income=txs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);
    const expense=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
    return{name,income,expense,balance:income-expense,count:txs.length};
  });
  const totInc=months.reduce((s,m)=>s+m.income,0),totExp=months.reduce((s,m)=>s+m.expense,0);
  // top individual tx — maior receita positiva
  const topInc=allTxs.filter(t=>t.type==='income').sort((a,b)=>b.amount-a.amount)[0];
  const topExp=allTxs.filter(t=>t.type==='expense').sort((a,b)=>b.amount-a.amount)[0];
  document.getElementById('ins-top-income').textContent=topInc?fmt(topInc.amount):'—';
  document.getElementById('ins-top-income-sub').textContent=topInc?topInc.description+' · '+new Date(topInc.date+'T12:00').toLocaleDateString('pt-BR'):'Nenhuma receita';
  document.getElementById('ins-top-expense').textContent=topExp?fmt(topExp.amount):'—';
  document.getElementById('ins-top-expense-sub').textContent=topExp?topExp.description+' · '+new Date(topExp.date+'T12:00').toLocaleDateString('pt-BR'):'Nenhuma despesa';
  // meses com movimentação de receita (inclusive negativos)
  const actMths=months.filter(m=>m.count>0&&allTxs.filter(t=>t.type==='income'&&new Date(t.date+'T12:00').getMonth()===months.indexOf(m)).length>0);
  const avgSave=actMths.length?actMths.reduce((s,m)=>s+(m.income!==0?(m.balance/Math.abs(m.income))*100:0),0)/actMths.length:0;
  document.getElementById('ins-save-rate').textContent=avgSave.toFixed(1)+'%';
  document.getElementById('ins-save-rate').style.color=avgSave>=0?'var(--green)':'var(--red)';
  document.getElementById('ins-save-sub').textContent='Média dos meses com receita';
  // ranking meses despesa
  const rankExp=months.filter(m=>m.expense>0).sort((a,b)=>b.expense-a.expense);
  const maxExp=rankExp[0]?.expense||1;
  document.getElementById('ins-rank-expense').innerHTML=rankExp.slice(0,6).map((m,i)=>`<div class="rank-row"><div class="rank-num">${i+1}</div><div style="font-size:13px;color:var(--text2);min-width:70px">${m.name.substring(0,3)}</div><div class="rank-bar-wrap"><div class="rank-bar" style="width:${(m.expense/maxExp*100).toFixed(0)}%;background:var(--red)"></div></div><div style="font-family:DM Mono;font-size:12px;color:var(--red);min-width:90px;text-align:right">${fmt(m.expense)}</div></div>`).join('')||'<div class="empty">Sem dados</div>';
  // ranking meses receita — inclui meses com receita negativa (count>0)
  const rankInc=months.filter(m=>{
    const mIdx=MONTHS.indexOf(m.name);
    return allTxs.filter(t=>t.type==='income'&&new Date(t.date+'T12:00').getMonth()===mIdx).length>0;
  }).sort((a,b)=>b.income-a.income);
  const maxIncAbs=Math.max(...rankInc.map(m=>Math.abs(m.income)),1);
  document.getElementById('ins-rank-income').innerHTML=rankInc.slice(0,6).map((m,i)=>{
    const color=m.income<0?'var(--amber)':'var(--green)';
    return`<div class="rank-row"><div class="rank-num">${i+1}</div><div style="font-size:13px;color:var(--text2);min-width:70px">${m.name.substring(0,3)}</div><div class="rank-bar-wrap"><div class="rank-bar" style="width:${(Math.abs(m.income)/maxIncAbs*100).toFixed(0)}%;background:${color}"></div></div><div style="font-family:DM Mono;font-size:12px;color:${color};min-width:90px;text-align:right">${m.income<0?'−':''}${fmt(m.income)}</div></div>`;
  }).join('')||'<div class="empty">Sem dados</div>';
  // % por categoria despesa
  const cats=loadCats();
  const catData=cats.expense.map(c=>({cat:c,total:allTxs.filter(t=>t.type==='expense'&&t.category===c.id).reduce((s,t)=>s+t.amount,0)})).filter(x=>x.total>0).sort((a,b)=>b.total-a.total);
  // categorias de receita com transações (inclusive negativas)
  const incCatData=cats.income.map(c=>{
    const catTxs=allTxs.filter(t=>t.type==='income'&&t.category===c.id);
    return{cat:c,total:catTxs.reduce((s,t)=>s+t.amount,0),count:catTxs.length};
  }).filter(x=>x.count>0).sort((a,b)=>b.total-a.total);
  chartInsCat=dc(chartInsCat);
  if(catData.length){
    chartInsCat=new Chart(document.getElementById('chart-ins-cat'),{type:'doughnut',data:{labels:catData.map(x=>x.cat.name),datasets:[{data:catData.map(x=>x.total),backgroundColor:catData.map(x=>x.cat.color),borderWidth:0,hoverOffset:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'55%',plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>' '+fmt(ctx.raw)+' ('+(totExp>0?(ctx.raw/totExp*100).toFixed(1):0)+'%)'}}}}});
  }
  document.getElementById('ins-cat-detail').innerHTML=catData.map(x=>{
    const pct=totExp>0?(x.total/totExp*100).toFixed(1):0;
    return`<div class="rank-row"><div style="width:10px;height:10px;border-radius:50%;background:${x.cat.color};flex-shrink:0"></div><div style="font-size:13px;flex:1">${x.cat.icon} ${x.cat.name}</div><div style="font-size:11px;color:var(--text3);margin-right:8px">${pct}%</div><div style="font-family:DM Mono;font-size:12px;color:var(--red)">${fmt(x.total)}</div></div>`;
  }).join('')+(incCatData.length?'<div style="padding:10px 0 4px;font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:.5px">Receitas por categoria</div>'+incCatData.map(x=>{
    const color=x.total<0?'var(--amber)':'var(--green)';
    return`<div class="rank-row"><div style="width:10px;height:10px;border-radius:50%;background:${x.cat.color};flex-shrink:0"></div><div style="font-size:13px;flex:1">${x.cat.icon} ${x.cat.name}</div><div style="font-family:DM Mono;font-size:12px;color:${color}">${x.total<0?'−':''}${fmt(x.total)}</div></div>`;
  }).join(''):'')||'<div class="empty">Sem dados no período</div>';
}

// ===================== CATEGORIAS =====================
function renderCategorias(){
  const month=parseInt(document.getElementById('cat-month')?.value??new Date().getMonth());
  const year=parseInt(document.getElementById('cat-year')?.value??new Date().getFullYear());
  const db=visibleDB();
  // month=-1 → todos os meses do ano selecionado
  const txs=month===-1
    ?db.transactions.filter(t=>new Date(t.date+'T12:00').getFullYear()===year)
    :filterTx(db.transactions,month,year);
  const cats=loadCats();
  function makeBars(type,cid){
    const grossTotal=txs.filter(t=>t.type===type).reduce((s,t)=>s+Math.abs(t.amount),0);
    const data=cats[type].map(c=>{
      const catTxs=txs.filter(t=>t.type===type&&t.category===c.id);
      const total=catTxs.reduce((s,t)=>s+t.amount,0);
      const count=catTxs.length;
      return{cat:c,total,count};
    // Para receita: mostra mesmo se total negativo (tem transações); para despesa: só se >0
    }).filter(x=>type==='income'?x.count>0:x.total>0).sort((a,b)=>b.total-a.total);
    document.getElementById(cid).innerHTML=!data.length?'<div class="empty" style="padding:16px">Sem dados</div>':data.map(x=>{
      const pct=grossTotal>0?Math.round((Math.abs(x.total)/grossTotal)*100):0;
      const amtColor=type==='income'&&x.total<0?'var(--amber)':'inherit';
      const sign=type==='income'&&x.total<0?'−':'';
      return`<div class="cat-row"><div class="cat-name-col">${x.cat.icon} ${x.cat.name}</div><div class="cat-bar-wrap"><div class="cat-bar" style="width:${pct}%;background:${x.cat.color}"></div></div><div class="cat-amt" style="color:${amtColor}">${sign}${fmt(x.total)}</div><div class="cat-pct">${pct}%</div></div>`;
    }).join('');
  }
  makeBars('expense','cat-expense-bars');makeBars('income','cat-income-bars');
  const expTotal=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
  const expData=cats.expense.map(c=>({cat:c,total:txs.filter(t=>t.type==='expense'&&t.category===c.id).reduce((s,t)=>s+t.amount,0)})).filter(x=>x.total>0).sort((a,b)=>b.total-a.total);
  chartCatPie=dc(chartCatPie);
  const catLeg=document.getElementById('cat-pie-legend');
  if(!expData.length){catLeg.innerHTML='';return}
  chartCatPie=new Chart(document.getElementById('chart-cat-pie'),{type:'doughnut',data:{labels:expData.map(x=>x.cat.name),datasets:[{data:expData.map(x=>x.total),backgroundColor:expData.map(x=>x.cat.color),borderWidth:0,hoverOffset:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'60%',plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>' '+fmt(ctx.raw)}}}}});
  catLeg.innerHTML=expData.map(x=>{const pct=expTotal>0?Math.round((x.total/expTotal)*100):0;return`<span class="legend-item"><span class="legend-dot" style="background:${x.cat.color}"></span><span style="color:var(--text2)">${x.cat.name}</span><span style="margin-left:auto;font-size:11px;color:var(--text3)">${pct}%</span></span>`}).join('');
  catLeg.style.cssText='display:flex;flex-direction:column;gap:7px;margin-top:14px';
}

// ===================== ANUAL =====================
function renderAnual(){
  const year=parseInt(document.getElementById('anual-year')?.value??new Date().getFullYear());
  const db=visibleDB();
  const allTxs=db.transactions.filter(t=>new Date(t.date+'T12:00').getFullYear()===year);
  const months=MONTHS.map((name,m)=>{const txs=filterTx(db.transactions,m,year);const income=txs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);const expense=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);return{name,income,expense,balance:income-expense,count:txs.length}});
  const totI=months.reduce((s,m)=>s+m.income,0),totE=months.reduce((s,m)=>s+m.expense,0),totB=totI-totE;
  document.getElementById('anual-income').textContent=fmt(totI);document.getElementById('anual-expense').textContent=fmt(totE);document.getElementById('anual-balance').textContent=fmt(totB);document.getElementById('anual-balance').style.color=totB>=0?'var(--green)':'var(--red)';
  chartAnual=dc(chartAnual);
  chartAnual=new Chart(document.getElementById('chart-anual'),{type:'bar',data:{labels:MONTHS_SHORT,datasets:[{label:'Receitas',data:months.map(x=>x.income),backgroundColor:'rgba(34,211,160,0.7)',borderRadius:4},{label:'Despesas',data:months.map(x=>x.expense),backgroundColor:'rgba(245,101,101,0.7)',borderRadius:4},{type:'line',label:'Saldo',data:months.map(x=>x.balance),borderColor:'rgba(96,165,250,0.9)',backgroundColor:'transparent',borderWidth:2,pointRadius:3,tension:0.3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}});
  document.getElementById('anual-tbody').innerHTML=months.map(m=>{
    const bc=m.balance>=0?'var(--green)':'var(--red)';
    const incColor=m.income<0?'var(--amber)':'var(--green)';
    const hasIncome=allTxs.filter(t=>t.type==='income'&&new Date(t.date+'T12:00').getFullYear()===year&&new Date(t.date+'T12:00').getMonth()===months.indexOf(m)).length>0;
    return`<tr style="border-bottom:1px solid var(--border);${!m.count?'opacity:.35':''}"><td style="padding:11px 18px;font-weight:500">${m.name}</td><td style="padding:11px 18px;text-align:right;color:${incColor};font-family:DM Mono;font-size:12px">${hasIncome?fmt(m.income):'—'}</td><td style="padding:11px 18px;text-align:right;color:var(--red);font-family:DM Mono;font-size:12px">${m.expense>0?fmt(m.expense):'—'}</td><td style="padding:11px 18px;text-align:right;color:${bc};font-family:DM Mono;font-size:12px">${m.count?fmt(m.balance):'—'}</td><td style="padding:11px 18px;text-align:right;color:var(--text2)">${m.count||'—'}</td></tr>`;
  }).join('');
  populateCatAnualSelect();renderCatAnual();
}
function populateCatAnualSelect(){
  const type=document.getElementById('cat-anual-type')?.value||'expense';
  const sel=document.getElementById('cat-anual-sel');if(!sel)return;
  sel.innerHTML=loadCats()[type].map(c=>`<option value="${c.id}">${c.icon} ${c.name}</option>`).join('');
}
function renderCatAnual(){
  const year=parseInt(document.getElementById('anual-year')?.value??new Date().getFullYear());
  const type=document.getElementById('cat-anual-type')?.value||'expense';
  const catId=document.getElementById('cat-anual-sel')?.value;if(!catId)return;
  const db=visibleDB(),cat=getCat(catId);
  const monthData=MONTHS.map((_,m)=>{const txs=filterTx(db.transactions,m,year).filter(t=>t.type===type&&t.category===catId);return{short:MONTHS_SHORT[m],name:MONTHS[m],total:txs.reduce((s,t)=>s+t.amount,0),count:txs.length}});
  const annual=monthData.reduce((s,m)=>s+m.total,0);
  const actv=monthData.filter(m=>m.total>0);
  const avg=actv.length?actv.reduce((s,m)=>s+m.total,0)/actv.length:0;
  chartCatAnual=dc(chartCatAnual);
  chartCatAnual=new Chart(document.getElementById('chart-cat-anual'),{type:'bar',data:{labels:monthData.map(x=>x.short),datasets:[{label:cat.name,data:monthData.map(x=>x.total),backgroundColor:monthData.map(x=>x.total>0?cat.color+'bb':'rgba(255,255,255,0.05)'),borderColor:cat.color,borderWidth:1,borderRadius:5},{type:'line',label:'Média',data:monthData.map(()=>avg),borderColor:'rgba(255,255,255,0.25)',borderDash:[4,4],borderWidth:1.5,pointRadius:0,backgroundColor:'transparent'}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>` ${fmt(ctx.raw)}`}}},scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(1)+'k'}}}}});
  const best=monthData.reduce((a,b)=>b.total>a.total?b:a,monthData[0]);
  document.getElementById('cat-anual-summary').innerHTML=[`<span style="background:var(--bg3);padding:5px 12px;border-radius:8px">📦 Total: <strong style="color:${cat.color}">${fmt(annual)}</strong></span>`,`<span style="background:var(--bg3);padding:5px 12px;border-radius:8px">📊 Média: <strong>${fmt(avg)}</strong></span>`,best.total>0?`<span style="background:var(--bg3);padding:5px 12px;border-radius:8px">🏆 Maior: <strong>${best.name} (${fmt(best.total)})</strong></span>`:'',`<span style="background:var(--bg3);padding:5px 12px;border-radius:8px">🔢 Lançamentos: <strong>${monthData.reduce((s,m)=>s+m.count,0)}</strong></span>`].join('');
}

// ===================== IMPORT =====================
let pendingImport=[];
function dragOver(e){e.preventDefault();document.getElementById('drop-zone').classList.add('dragover')}
function dragLeave(){document.getElementById('drop-zone').classList.remove('dragover')}
function dropFile(e){e.preventDefault();dragLeave();const f=e.dataTransfer.files[0];if(f)handleFile(f)}
function handleFile(file){
  if(!file)return;
  const reader=new FileReader();
  reader.onload=(e)=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'binary',cellDates:true});
      const sn=wb.SheetNames.find(s=>/gastos?/i.test(s))||wb.SheetNames[0];
      const ws=wb.Sheets[sn];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
      let hi=0;
      for(let i=0;i<Math.min(10,rows.length);i++){const r=rows[i];if(r.some(c=>c&&/^dia$|^data$/i.test(String(c).trim()))&&r.some(c=>c&&/^valor$/i.test(String(c).trim()))){hi=i;break}}
      const hdr=rows[hi].map(c=>String(c||'').toLowerCase().trim());
      let cDia=-1,cTipo=-1,cDept=-1,cValor=-1,cDesc=-1;
      hdr.forEach((h,i)=>{if(/^dia$|^data$/.test(h))cDia=i;else if(/^valor$/.test(h))cValor=i;else if(/departamento|categoria/.test(h))cDept=i;else if(/descri/.test(h))cDesc=i;else if(/^tipo$/.test(h))cTipo=i;else if((h===''||h==='unnamed')&&cTipo===-1)cTipo=i});
      if(cDia<0)cDia=0;if(cTipo<0)cTipo=1;if(cDept<0)cDept=2;if(cValor<0)cValor=3;if(cDesc<0)cDesc=4;
      const dataRows=rows.slice(hi+1),parsed=[],skipped=[],newCatMap={};
      const cats=loadCats(),allEx=[...cats.expense,...cats.income];
      const norm=s=>String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim();
      dataRows.forEach((row,idx)=>{
        if(!row){skipped.push(idx);return}
        const rawDate=row[cDia],rawTipo=String(row[cTipo]||'').trim(),rawDept=String(row[cDept]||'').trim(),rawValor=row[cValor],rawDesc=String(row[cDesc]||'').trim();
        if(!rawDate||rawValor==null||isNaN(Number(rawValor))){skipped.push(idx);return}
        const rawNum=Number(rawValor),amount=rawNum;if(amount===0){skipped.push(idx);return}
        let dateStr='';
        if(rawDate instanceof Date&&!isNaN(rawDate)){dateStr=`${rawDate.getFullYear()}-${String(rawDate.getMonth()+1).padStart(2,'0')}-${String(rawDate.getDate()).padStart(2,'0')}`}
        else if(typeof rawDate==='number'){const d=new Date(Math.round((rawDate-25569)*86400*1000));if(!isNaN(d))dateStr=d.toISOString().split('T')[0]}
        else{const d=new Date(String(rawDate));if(!isNaN(d))dateStr=d.toISOString().split('T')[0]}
        if(!dateStr){skipped.push(idx);return}
        const tipoN=norm(rawTipo);let type='expense';
        if(/recebid|entrada|receita|income/.test(tipoN))type='income';
        const deptN=norm(rawDept);
        // Doação é sempre despesa, independente do tipo marcado na planilha
        if(/doa[cç][aã]o/.test(deptN)) type='expense';
        let catId=null;
        // Apenas correspondência EXATA (normalizada) — evita confundir categorias parecidas
        // como "Receber" vs "Recebidos" ou "Presente" vs "Presentes"
        for(const c of allEx){if(deptN===norm(c.name)){catId=c.id;break}}
        if(!catId){catId='imp_'+slugify(rawDept);if(!newCatMap[catId])newCatMap[catId]={id:catId,name:rawDept.trim(),type,icon:'🏷️',color:randomColor()}}
        const uid=`${dateStr}_${amount}_${catId}_${(rawDesc||rawDept).slice(0,12).replace(/\s/g,'_')}`;
        parsed.push({id:uid,type,description:rawDesc||rawDept,amount,date:dateStr,category:catId});
      });
      pendingImport=parsed;
      document.getElementById('import-filename').textContent=file.name;
      document.getElementById('import-fileinfo').textContent=`Aba: "${sn}" • ${dataRows.length} linhas • Dia=${cDia} Tipo=${cTipo} Dept=${cDept} Valor=${cValor}`;
      document.getElementById('imp-total').textContent=parsed.length+skipped.length;
      document.getElementById('imp-income').textContent=parsed.filter(t=>t.type==='income').length;
      document.getElementById('imp-expense').textContent=parsed.filter(t=>t.type==='expense').length;
      document.getElementById('imp-skip').textContent=skipped.length;
      document.getElementById('import-tbody').innerHTML=parsed.slice(0,20).map(t=>{const d=new Date(t.date+'T12:00'),color=t.type==='income'?'var(--green)':'var(--red)',sign=t.amount>=0?(t.type==='income'?'+':'-'):'-';return`<tr><td>${d.toLocaleDateString('pt-BR')}</td><td style="color:${color}">${t.type==='income'?'Receita':'Despesa'}</td><td>${getCat(t.category).name||t.category}</td><td style="color:${color};font-family:DM Mono">${sign} ${fmt(t.amount)}</td><td>${t.description}</td></tr>`}).join('');
      const newCats=Object.values(newCatMap);
      document.getElementById('import-new-cats').innerHTML=newCats.length?newCats.map(c=>`<span style="display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;background:var(--bg3);border:1px solid var(--border);font-size:12px"><span style="width:7px;height:7px;border-radius:50%;background:${c.color}"></span>${c.icon} ${c.name}</span>`).join(''):'<span style="font-size:12px;color:var(--text3)">Todas já existem</span>';
      const warn=document.getElementById('import-warn');
      if(!parsed.length){warn.style.display='block';warn.textContent='⚠ Nenhuma linha válida. Verifique o formato.'}else warn.style.display='none';
      document.getElementById('import-preview').style.display='block';
      if(document.getElementById('imp-owner'))document.getElementById('imp-owner').value=_currentProfile;
    }catch(err){alert('Erro: '+err.message)}
  };
  reader.readAsBinaryString(file);
}
async function confirmImport(){
  if(!pendingImport.length){alert('Nenhum dado para importar.');return}
  const btn=document.getElementById('btn-confirm-import');btn.textContent='⏳...';btn.disabled=true;
  const skipDup=document.getElementById('imp-skip-dup').checked;
  const importOwner=document.getElementById('imp-owner')?.value||_currentProfile;
  const cats=loadCats(),allIds=new Set([...cats.expense,...cats.income].map(c=>c.id));
  pendingImport.forEach(t=>{t.owner=importOwner;if(!allIds.has(t.category)){cats[t.type]=cats[t.type]||[];cats[t.type].push({id:t.category,name:t.category.replace('imp_','').replace(/_/g,' '),icon:'🏷️',color:randomColor(),builtin:false});allIds.add(t.category)}});
  await saveCats(cats);
  const db=loadDB(),existingIds=skipDup?new Set(db.transactions.map(t=>t.id)):new Set();
  let added=0,dupes=0;
  pendingImport.forEach(t=>{if(skipDup&&existingIds.has(t.id)){dupes++;return}db.transactions.push(t);existingIds.add(t.id);added++});
  db.transactions.sort((a,b)=>new Date(b.date)-new Date(a.date));
  await saveDBAndSync(db);
  const history=loadImports();history.unshift({date:new Date().toISOString(),added,dupes,total:pendingImport.length});
  await saveImports(history.slice(0,20));
  btn.textContent='✓ Importar';btn.disabled=false;
  alert(`✓ ${added} transações adicionadas${dupes?' ('+dupes+' duplicatas ignoradas)':''}.`);
  resetImport();populateAllSelects();refreshAll();renderImportHistory();
}
function resetImport(){pendingImport=[];document.getElementById('import-preview').style.display='none';document.getElementById('file-input').value=''}
function renderImportHistory(){
  const h=loadImports(),el=document.getElementById('import-history-list');
  if(!h.length){el.innerHTML='<div class="empty"><div class="empty-icon">📋</div>Nenhuma importação ainda</div>';return}
  el.innerHTML=h.map(x=>{const d=new Date(x.date);return`<div class="tx-item"><div class="tx-icon" style="background:var(--green-bg)">📥</div><div class="tx-info"><div class="tx-name">${x.added} transações importadas</div><div class="tx-cat">${x.dupes||0} duplicatas ignoradas</div></div><div class="tx-date">${d.toLocaleDateString('pt-BR')} ${d.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})}</div></div>`}).join('');
}

// ===================== EXPORT =====================
function renderExportSection(){
  populateMonthSelect('exp-month',new Date().getMonth());populateYearSelect('exp-year',new Date().getFullYear());
  const sel=document.getElementById('exp-cat');
  sel.innerHTML='<option value="all">Todas</option>'+allCats().map(c=>`<option value="${c.id}">${c.icon} ${c.name}</option>`).join('');
  if(document.getElementById('exp-owner')&&!document.getElementById('exp-owner')._init){document.getElementById('exp-owner')._init=true}
  updateExportPreview();
  ['exp-month','exp-year','exp-type','exp-cat','exp-owner'].forEach(id=>{const el=document.getElementById(id);if(el)el.onchange=updateExportPreview});
}
function getExportTxs(){
  const month=parseInt(document.getElementById('exp-month')?.value??-1);
  const year=parseInt(document.getElementById('exp-year')?.value??new Date().getFullYear());
  const type=document.getElementById('exp-type')?.value||'all';
  const catId=document.getElementById('exp-cat')?.value||'all';
  const owner=document.getElementById('exp-owner')?.value||'ambos';
  const db=loadDB();
  let txs=month===-1?db.transactions.filter(t=>new Date(t.date+'T12:00').getFullYear()===year):filterTx(db.transactions,month,year);
  if(type!=='all')txs=txs.filter(t=>t.type===type);
  if(catId!=='all')txs=txs.filter(t=>t.category===catId);
  if(owner!=='ambos')txs=txs.filter(t=>t.owner===owner);
  return txs;
}
function updateExportPreview(){
  const txs=getExportTxs();
  const cnt=document.getElementById('exp-preview-count');if(cnt)cnt.textContent=txs.length+' transação(ões)';
  const tb=document.getElementById('exp-preview-tbody');if(!tb)return;
  if(!txs.length){tb.innerHTML='<tr><td colspan="6" style="text-align:center;padding:20px;color:var(--text3)">Sem transações no período</td></tr>';return}
  tb.innerHTML=txs.slice(0,50).map(t=>{const c=getCat(t.category),d=new Date(t.date+'T12:00'),color=t.type==='income'?'var(--green)':'var(--red)',sign=t.type==='income'?'+':'-';const ownerLbl=profileName(t.owner)+(t.shared?' (🤝'+t.splitPct+'%)':'');return`<tr><td>${d.toLocaleDateString('pt-BR')}</td><td style="color:${color}">${t.type==='income'?'Receita':'Despesa'}</td><td>${ownerLbl}</td><td>${c.icon} ${c.name}</td><td style="color:${color};font-family:DM Mono">${sign} ${fmt(t.amount)}</td><td>${t.description}</td></tr>`}).join('')+(txs.length>50?`<tr><td colspan="6" style="text-align:center;padding:8px;color:var(--text3);font-size:12px">... e mais ${txs.length-50} transações</td></tr>`:'');
}
function exportExcel(){
  const txs=getExportTxs();if(!txs.length){alert('Nenhuma transação para exportar.');return}
  const rows=[['Data','Tipo','Dono','Compartilhado','% desta parte','Total original (se compartilhado)','Departamento','Valor','Descrição','Mês','Ano']];
  txs.forEach(t=>{const d=new Date(t.date+'T12:00'),c=getCat(t.category);rows.push([t.date,t.type==='income'?'Recebido':'Gasto',profileName(t.owner),t.shared?'Sim':'Não',t.shared?t.splitPct:'',t.shared?t.sharedTotal:'',c.name,t.type==='expense'?-t.amount:t.amount,t.description,d.getMonth()+1,d.getFullYear()])});
  const wb=XLSX.utils.book_new(),ws=XLSX.utils.aoa_to_sheet(rows);
  ws['!cols']=[{wch:12},{wch:10},{wch:10},{wch:12},{wch:12},{wch:16},{wch:20},{wch:12},{wch:30},{wch:6},{wch:6}];
  const range=XLSX.utils.decode_range(ws['!ref']);
  for(let r=1;r<=range.e.r;r++){const dc2=ws[XLSX.utils.encode_cell({r,c:0})];if(dc2)dc2.z='DD/MM/YYYY';const vc=ws[XLSX.utils.encode_cell({r,c:7})];if(vc){vc.t='n';vc.z='#,##0.00'}}
  XLSX.utils.book_append_sheet(wb,ws,'Gastos');
  const year=parseInt(document.getElementById('exp-year')?.value??new Date().getFullYear());
  const sRows=[['Mês','Receitas','Despesas','Saldo']];
  MONTHS.forEach((name,m)=>{const mt=filterTx(txs,m,year);const inc=mt.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);const exp=mt.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);if(inc||exp)sRows.push([name,inc,-exp,inc-exp])});
  if(sRows.length>1){const ws2=XLSX.utils.aoa_to_sheet(sRows);ws2['!cols']=[{wch:14},{wch:14},{wch:14},{wch:14}];XLSX.utils.book_append_sheet(wb,ws2,'Resumo Mensal')}
  const month=parseInt(document.getElementById('exp-month')?.value??-1);
  const ownerLbl=document.getElementById('exp-owner')?.value||'ambos';
  XLSX.writeFile(wb,`FinTrack_${ownerLbl}_${year}_${month===-1?'Anual':MONTHS_SHORT[month]}.xlsx`);
}

// ===================== CAT MANAGER =====================
function renderCatManager(){
  const cats=loadCats();
  const db=loadDB();
  // compute total per category
  const totals={};
  db.transactions.forEach(t=>{totals[t.category]=(totals[t.category]||0)+t.amount});

  ['expense','income'].forEach(type=>{
    const el=document.getElementById('gcat-'+type+'-list');
    if(!el)return;
    el.innerHTML=cats[type].map(c=>{
      const total=totals[c.id]||0;
      const inUse=db.transactions.some(t=>t.category===c.id);
      const totalColor=total<0?'var(--amber)':total>0?'var(--green)':'var(--text3)';
      const totalBadge=inUse?`<span class="cat-total-badge" style="color:${totalColor}">${total<0?'−':''}${fmt(total)}</span>`:'';
      const editBtn=`<button class="cat-item-edit-btn" onclick="toggleCatEdit('${type}','${c.id}')" title="Editar">✏️</button>`;
      const delBtn=`<button class="cat-item-del" onclick="deleteCat('${type}','${c.id}')" title="Excluir">✕</button>`;
      return`<div>
        <div class="cat-item-row" id="crow-${c.id}">
          <div class="cat-color-dot" style="background:${c.color}"></div>
          <div style="font-size:16px">${c.icon}</div>
          <div class="cat-item-name">${c.name}</div>
          ${totalBadge}
          ${c.builtin?'<span class="cat-item-badge">padrão</span>':''}
          ${editBtn}
          ${delBtn}
        </div>
        <div id="cedit-${c.id}" class="cat-edit-inline" style="display:none">
          <input id="cename-${c.id}" class="form-input" value="${c.name}" placeholder="Nome" style="flex:1;min-width:80px;height:32px;padding:4px 8px;font-size:12px">
          <input type="color" id="cecolor-${c.id}" value="${c.color}" class="color-picker" style="width:32px;height:32px">
          <input id="ceicon-${c.id}" class="form-input" value="${c.icon}" placeholder="🏷️" style="width:40px;text-align:center;height:32px;padding:4px;font-size:14px">
          <button class="btn btn-primary" onclick="saveCatEdit('${type}','${c.id}')" style="padding:4px 10px;font-size:12px">✓</button>
          <button class="btn" onclick="toggleCatEdit('${type}','${c.id}')" style="padding:4px 10px;font-size:12px">✕</button>
        </div>
      </div>`;
    }).join('')||'<div style="color:var(--text3);font-size:13px;padding:8px 0">Nenhuma</div>';
  });
}

function toggleCatEdit(type,id){
  const el=document.getElementById('cedit-'+id);
  if(!el)return;
  el.style.display=el.style.display==='none'?'flex':'none';
}

async function saveCatEdit(type,id){
  const name=document.getElementById('cename-'+id)?.value.trim();
  const color=document.getElementById('cecolor-'+id)?.value;
  const icon=document.getElementById('ceicon-'+id)?.value.trim()||'📦';
  if(!name){alert('Nome não pode ser vazio.');return}
  const cats=loadCats();
  const idx=cats[type].findIndex(c=>c.id===id);
  if(idx>=0){cats[type][idx]={...cats[type][idx],name,color,icon};}
  await saveCats(cats);
  renderCatManager();
}
async function addCat(type){
  const nE=document.getElementById(type==='expense'?'new-exp-name':'new-inc-name');
  const cE=document.getElementById(type==='expense'?'new-exp-color':'new-inc-color');
  const iE=document.getElementById(type==='expense'?'new-exp-icon':'new-inc-icon');
  const name=nE.value.trim();if(!name){alert('Digite um nome.');return}
  const cats=loadCats();
  cats[type].push({id:slugify(name)+'_'+Date.now().toString(36),name,icon:iE.value.trim()||(type==='expense'?'📦':'💰'),color:cE.value,builtin:false});
  await saveCats(cats);nE.value='';iE.value='';renderCatManager();
}
async function deleteCat(type,id){
  const cats=loadCats(),cat=cats[type].find(c=>c.id===id);
  const inUse=loadDB().transactions.some(t=>t.category===id);
  if(inUse){
    if(!confirm(`A categoria "${cat?.name}" tem transações registradas.\nAs transações existentes manterão a referência à categoria, mas ela não aparecerá mais nos filtros.\n\nExcluir mesmo assim?`))return;
  } else if(cat?.builtin){
    if(!confirm(`"${cat?.name}" é uma categoria padrão. Deseja excluí-la mesmo assim?`))return;
  }
  cats[type]=cats[type].filter(c=>c.id!==id);
  await saveCats(cats);renderCatManager();
}

// ===================== CLOUD =====================
function renderCloudSection(){
  const url=_fbUrl;
  document.getElementById('firebase-url').value=url||'';
  const badge=document.getElementById('sync-status-badge');
  if(url){badge.className='sync-status sync-ok';badge.innerHTML='<span class="sync-dot ok"></span>Firebase conectado';document.getElementById('cloud-actions').style.display='block';document.getElementById('cloud-url-display').textContent=url}
  else{badge.className='sync-status sync-local';badge.innerHTML='<span class="sync-dot local"></span>Armazenamento local';document.getElementById('cloud-actions').style.display='none'}
}
async function connectFirebase(){
  let url=document.getElementById('firebase-url').value.trim();if(!url){alert('Cole a URL.');return}
  if(!url.startsWith('https://'))url='https://'+url;if(url.endsWith('/'))url=url.slice(0,-1);
  try{const r=await fetch(url+'/.json?shallow=true');if(!r.ok)throw new Error('HTTP '+r.status);_fbUrl=url;await store.set(FB_KEY,url);logCloud('✓ Conectado');renderCloudSection()}
  catch(err){alert('Não foi possível conectar.\n'+err.message)}
}
function disconnectFirebase(){if(!confirm('Desconectar?'))return;_fbUrl='';store.set(FB_KEY,'');renderCloudSection()}
async function syncToCloud(){
  if(!_fbUrl){alert('Configure o Firebase primeiro.');return}
  try{logCloud('⏳ Enviando...');const payload={db:JSON.stringify(_db),cats:JSON.stringify(_cats),imports:JSON.stringify(_imports),updated:new Date().toISOString()};const r=await fetch(_fbUrl+'/fintrack.json',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});if(!r.ok)throw new Error('HTTP '+r.status);logCloud('✓ Enviado em '+new Date().toLocaleTimeString('pt-BR'))}
  catch(err){logCloud('✗ Erro: '+err.message)}
}
async function syncFromCloud(){
  if(!_fbUrl){alert('Configure o Firebase primeiro.');return}
  try{logCloud('⏳ Baixando...');const r=await fetch(_fbUrl+'/fintrack.json');if(!r.ok)throw new Error('HTTP '+r.status);const data=await r.json();if(!data){logCloud('⚠ Sem dados.');return}if(data.db)await saveDB(JSON.parse(data.db));if(data.cats)await saveCats(JSON.parse(data.cats));if(data.imports)await saveImports(JSON.parse(data.imports));populateAllSelects();refreshAll();logCloud('✓ Baixado em '+new Date().toLocaleTimeString('pt-BR')+(data.updated?' (salvo: '+new Date(data.updated).toLocaleString('pt-BR')+')':''))}
  catch(err){logCloud('✗ Erro: '+err.message)}
}
function logCloud(msg){const el=document.getElementById('cloud-log');if(el)el.innerHTML=msg+'<br>'+el.innerHTML}

// ===================== RESET =====================
function renderResetSection(){
  const db=loadDB(),cats=loadCats(),imps=loadImports();
  const cExp=cats.expense.filter(c=>!c.builtin).length,cInc=cats.income.filter(c=>!c.builtin).length;
  const tc=db.transactions.length,ti=db.transactions.filter(t=>t.type==='income').length,te=db.transactions.filter(t=>t.type==='expense').length;
  const txEl=document.getElementById('reset-tx-count');if(txEl)txEl.innerHTML=tc?`<strong style="color:var(--text)">${tc}</strong> transações — <span style="color:var(--green)">${ti} receitas</span> · <span style="color:var(--red)">${te} despesas</span>`:'Nenhuma transação.';
  const catEl=document.getElementById('reset-cats-count');if(catEl)catEl.innerHTML=(cExp+cInc)?`<strong style="color:var(--text)">${cExp+cInc}</strong> personalizadas — ${cExp} despesa · ${cInc} receita`:'Nenhuma personalizada.';
  const impEl=document.getElementById('reset-imp-count');if(impEl)impEl.innerHTML=imps.length?`<strong style="color:var(--text)">${imps.length}</strong> importação(ões)`:'Histórico vazio.';
  const cb=document.getElementById('confirm-reset-all');if(cb)cb.checked=false;
}
async function resetTransactions(){
  const db=loadDB();if(!db.transactions.length){alert('Não há transações.');return}
  if(!confirm(`Apagar ${db.transactions.length} transações?`))return;
  await saveDBAndSync({transactions:[]});populateAllSelects();renderResetSection();refreshAll();alert('✓ Transações apagadas.');
}
async function resetCustomCats(){
  const cats=loadCats(),total=cats.expense.filter(c=>!c.builtin).length+cats.income.filter(c=>!c.builtin).length;
  if(!total){alert('Nenhuma categoria personalizada.');return}
  if(!confirm(`Remover ${total} categoria(s)?`))return;
  cats.expense=cats.expense.filter(c=>c.builtin);cats.income=cats.income.filter(c=>c.builtin);
  await saveCats(cats);renderResetSection();renderCatManager();alert('✓ Categorias removidas.');
}
async function resetImportHistory(){
  if(!loadImports().length){alert('Histórico já vazio.');return}
  if(!confirm('Limpar histórico?'))return;await saveImports([]);renderResetSection();alert('✓ Histórico limpo.');
}
async function resetAll(){
  const cb=document.getElementById('confirm-reset-all');if(!cb?.checked){alert('Marque a confirmação.');return}
  if(!confirm('Apagar TUDO? Irreversível.'))return;
  await saveDBAndSync({transactions:[]});await saveCats(JSON.parse(JSON.stringify(DEFAULT_CATS)));await saveImports([]);
  populateAllSelects();renderResetSection();refreshAll();alert('✓ Tudo apagado.');
}

// ===================== REFRESH =====================
function refreshAll(){
  const id=document.querySelector('.section.active')?.id;
  if(id==='sec-dashboard')updateDashboard();
  if(id==='sec-transacoes')renderTxList();
  if(id==='sec-categorias')renderCategorias();
  if(id==='sec-anual')renderAnual();
  if(id==='sec-evolucao')renderEvolucao();
  if(id==='sec-insights')renderInsights();
}
function populateAllSelects(){
  const now=new Date(),cy=now.getFullYear(),cm=now.getMonth();
  populateMonthSelect('dash-month',cm);populateYearSelect('dash-year',cy);
  populateMonthSelect('tx-month',cm);populateYearSelect('tx-year',cy);
  // cat-month: always keep "Todos os meses" as first option
  const catMonthSel=document.getElementById('cat-month');
  if(catMonthSel){
    catMonthSel.innerHTML='<option value="-1">Todos os meses</option>'+
      MONTHS.map((m,i)=>`<option value="${i}"${i===cm?' selected':''}>${m}</option>`).join('');
  }
  populateYearSelect('cat-year',cy);
  populateYearSelect('anual-year',cy);populateYearSelect('exp-year',cy);
  populateYearSelect('ev-year',cy);populateYearSelect('ins-year',cy);
  populateMonthSelect('cmp-month',cm);populateYearSelect('cmp-year',cy);
}

// ===================== COMPARATIVO =====================
let chartComparativo=null;

function renderComparativo(){
  const mode=document.getElementById('cmp-mode')?.value||'ytd';
  const year=parseInt(document.getElementById('cmp-year')?.value??new Date().getFullYear());
  const prevYear=year-1;
  const cmpMonthSel=document.getElementById('cmp-month');
  if(cmpMonthSel) cmpMonthSel.style.display=mode==='month'?'inline-block':'none';
  const selMonth=mode==='month'?parseInt(cmpMonthSel?.value??new Date().getMonth()):null;
  const db=visibleDB();

  // Define which months to include
  const now=new Date();
  const maxMonth=mode==='ytd'?(year===now.getFullYear()?now.getMonth():11):11;

  function getTxsForPeriod(y,months){
    return db.transactions.filter(t=>{
      const d=new Date(t.date+'T12:00');
      return d.getFullYear()===y&&months.includes(d.getMonth());
    });
  }

  let monthRange=[];
  if(mode==='month') monthRange=[selMonth];
  else if(mode==='ytd') monthRange=Array.from({length:maxMonth+1},(_,i)=>i);
  else monthRange=Array.from({length:12},(_,i)=>i);

  const curTxs=getTxsForPeriod(year,monthRange);
  const prevTxs=getTxsForPeriod(prevYear,monthRange);

  const curInc=curTxs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);
  const curExp=curTxs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
  const curBal=curInc-curExp;
  const prevInc=prevTxs.filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0);
  const prevExp=prevTxs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
  const prevBal=prevInc-prevExp;

  // subtitle
  const modeLabel=mode==='ytd'?`Jan–${MONTHS_SHORT[maxMonth]}`:(mode==='month'?MONTHS[selMonth]:'Ano completo');
  document.getElementById('cmp-subtitle').textContent=`${modeLabel} ${year} vs ${modeLabel} ${prevYear}`;

  // update labels
  ['','2','3'].forEach(s=>{
    const cl=document.getElementById('cmp-year-cur-label'+s);
    const pl=document.getElementById('cmp-year-prev-label'+s);
    if(cl)cl.textContent=year;if(pl)pl.textContent=prevYear;
  });

  // fill cards
  function setBadge(id,cur,prev,invertGood){
    const el=document.getElementById(id);if(!el)return;
    if(!prev){el.innerHTML='<span class="cmp-badge cmp-flat">—</span>';return}
    const pct=((cur-prev)/Math.abs(prev)*100);
    const up=cur>prev;
    // for expenses: up is bad; for income/balance: up is good
    const good=invertGood?!up:up;
    const cls=Math.abs(pct)<0.5?'cmp-flat':good?'cmp-up':'cmp-down';
    const arrow=Math.abs(pct)<0.5?'→':up?'↑':'↓';
    el.innerHTML=`<span class="cmp-badge ${cls}">${arrow} ${Math.abs(pct).toFixed(1)}%</span>`;
  }

  document.getElementById('cmp-inc-cur').textContent=fmt(curInc);
  document.getElementById('cmp-inc-prev').textContent=fmt(prevInc);
  setBadge('cmp-inc-badge',curInc,prevInc,false);

  document.getElementById('cmp-exp-cur').textContent=fmt(curExp);
  document.getElementById('cmp-exp-prev').textContent=fmt(prevExp);
  setBadge('cmp-exp-badge',curExp,prevExp,true);

  document.getElementById('cmp-bal-cur').textContent=fmt(curBal);
  document.getElementById('cmp-bal-cur').style.color=curBal>=0?'var(--green)':'var(--red)';
  document.getElementById('cmp-bal-prev').textContent=fmt(prevBal);
  setBadge('cmp-bal-badge',curBal,prevBal,false);

  // legend labels
  document.getElementById('cmp-legend-cur').textContent='Despesa '+year;
  document.getElementById('cmp-legend-prev').textContent='Despesa '+prevYear;
  document.getElementById('cmp-legend-inc-cur').textContent='Receita '+year;
  document.getElementById('cmp-legend-inc-prev').textContent='Receita '+prevYear;

  // Chart: mês a mês dentro do período
  const labels=mode==='month'?[MONTHS[selMonth]]:monthRange.map(m=>MONTHS_SHORT[m]);
  const curExpByMonth=monthRange.map(m=>getTxsForPeriod(year,[m]).filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0));
  const prevExpByMonth=monthRange.map(m=>getTxsForPeriod(prevYear,[m]).filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0));
  const curIncByMonth=monthRange.map(m=>getTxsForPeriod(year,[m]).filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0));
  const prevIncByMonth=monthRange.map(m=>getTxsForPeriod(prevYear,[m]).filter(t=>t.type==='income').reduce((s,t)=>s+t.amount,0));

  chartComparativo=dc(chartComparativo);
  chartComparativo=new Chart(document.getElementById('chart-comparativo'),{
    type:'bar',
    data:{
      labels,
      datasets:[
        {label:'Despesa '+year,data:curExpByMonth,backgroundColor:'rgba(245,101,101,0.8)',borderRadius:3},
        {label:'Despesa '+prevYear,data:prevExpByMonth,backgroundColor:'rgba(245,101,101,0.25)',borderRadius:3},
        {type:'line',label:'Receita '+year,data:curIncByMonth,borderColor:'rgba(34,211,160,0.9)',backgroundColor:'transparent',borderWidth:2,pointRadius:3,tension:0.3},
        {type:'line',label:'Receita '+prevYear,data:prevIncByMonth,borderColor:'rgba(34,211,160,0.3)',backgroundColor:'transparent',borderWidth:1.5,borderDash:[4,3],pointRadius:2,tension:0.3},
      ]
    },
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{x:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8'}},
              y:{grid:{color:'rgba(255,255,255,0.05)'},ticks:{color:'#8b93a8',callback:v=>'R$'+(v/1000).toFixed(0)+'k'}}}}
  });

  // Category comparison
  const cats=loadCats();
  const allCatIds=[...cats.expense,...cats.income].map(c=>c.id);
  const catRows=cats.expense.map(c=>{
    const cur=curTxs.filter(t=>t.type==='expense'&&t.category===c.id).reduce((s,t)=>s+t.amount,0);
    const prev=prevTxs.filter(t=>t.type==='expense'&&t.category===c.id).reduce((s,t)=>s+t.amount,0);
    return{cat:c,cur,prev,diff:cur-prev,pct:prev>0?((cur-prev)/prev*100):null};
  }).filter(x=>x.cur>0||x.prev>0).sort((a,b)=>b.cur-a.cur);

  const maxCat=Math.max(...catRows.map(r=>Math.max(r.cur,r.prev)),1);
  document.getElementById('cmp-cat-list').innerHTML=catRows.map(r=>{
    const pctCur=Math.round(r.cur/maxCat*100);
    const pctPrev=Math.round(r.prev/maxCat*100);
    const diffColor=r.diff>0?'var(--red)':r.diff<0?'var(--green)':'var(--text3)';
    const diffSign=r.diff>0?'+':'';
    return`<div style="padding:10px 20px;border-bottom:1px solid var(--border)">
      <div style="display:flex;justify-content:space-between;margin-bottom:5px">
        <span style="font-size:13px">${r.cat.icon} ${r.cat.name}</span>
        <span style="font-size:11px;color:${diffColor};font-family:DM Mono">${diffSign}${fmt(r.diff)}</span>
      </div>
      <div style="display:flex;gap:6px;align-items:center;margin-bottom:3px">
        <span style="font-size:10px;color:var(--text3);width:28px">${year}</span>
        <div style="flex:1;height:5px;background:var(--bg3);border-radius:3px;overflow:hidden"><div style="width:${pctCur}%;height:100%;background:${r.cat.color};border-radius:3px"></div></div>
        <span style="font-size:11px;font-family:DM Mono;color:var(--red);min-width:80px;text-align:right">${r.cur>0?fmt(r.cur):'—'}</span>
      </div>
      <div style="display:flex;gap:6px;align-items:center">
        <span style="font-size:10px;color:var(--text3);width:28px">${prevYear}</span>
        <div style="flex:1;height:5px;background:var(--bg3);border-radius:3px;overflow:hidden"><div style="width:${pctPrev}%;height:100%;background:${r.cat.color};border-radius:3px;opacity:.35"></div></div>
        <span style="font-size:11px;font-family:DM Mono;color:var(--text2);min-width:80px;text-align:right">${r.prev>0?fmt(r.prev):'—'}</span>
      </div>
    </div>`;
  }).join('')||'<div class="empty">Sem dados para comparar</div>';

  // Maiores variações
  const variations=catRows.filter(r=>r.prev>0&&r.cur>0&&r.pct!==null).sort((a,b)=>Math.abs(b.pct)-Math.abs(a.pct)).slice(0,8);
  document.getElementById('cmp-variations').innerHTML=variations.map(r=>{
    const up=r.diff>0;
    const cls=up?'cmp-down':'cmp-up'; // expenses: up is bad
    return`<div class="rank-row">
      <div style="font-size:16px">${r.cat.icon}</div>
      <div style="flex:1;min-width:0"><div style="font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r.cat.name}</div><div style="font-size:11px;color:var(--text2)">${fmt(r.prev)} → ${fmt(r.cur)}</div></div>
      <span class="cmp-badge ${cls}">${up?'↑':'↓'} ${Math.abs(r.pct).toFixed(1)}%</span>
    </div>`;
  }).join('')||'<div class="empty">Sem variações para exibir</div>';
}

// ===================== PATRIMÔNIO =====================
let chartPatrimonio=null;
let _patData=null; // cache da projeção

// Categorias excluídas da projeção de receita (por nome normalizado)
const PAT_EXCLUDED_PATTERNS=[
  /rendimento/,/investimento/,/freelance/,/bonus/,/bônus/,
  /projeto/,/receber/,/recebido/,/doacao/,/doação/
];

function isExcludedFromProjection(catName){
  const n=catName.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  return PAT_EXCLUDED_PATTERNS.some(p=>p.test(n));
}

// Busca série BCB com fallback via proxy CORS
async function fetchBCBSerie(serie){
  // Tenta direto primeiro (funciona em localhost e alguns hosts)
  const direct=`https://api.bcb.gov.br/dados/serie/bcdata.sgs.${serie}/dados/ultimos/12?formato=json`;
  // Fallback via proxy CORS público
  const proxy=`https://api.allorigins.win/get?url=${encodeURIComponent(direct)}`;

  // Tenta direto
  try{
    const r=await fetch(direct,{signal:AbortSignal.timeout?AbortSignal.timeout(6000):undefined});
    if(r.ok){
      const data=await r.json();
      if(Array.isArray(data)&&data.length){
        return parseFloat(String(data[data.length-1].valor).replace(',','.'));
      }
    }
  }catch{}

  // Fallback proxy
  try{
    const r=await fetch(proxy);
    if(r.ok){
      const wrapper=await r.json();
      if(wrapper.contents){
        const data=JSON.parse(wrapper.contents);
        if(Array.isArray(data)&&data.length){
          return parseFloat(String(data[data.length-1].valor).replace(',','.'));
        }
      }
    }
  }catch{}

  return null;
}

async function fetchFocusSelic(){
  // Série 4390 = Meta SELIC definida pelo Copom (% a.a.) - mais estável
  // Série 432  = SELIC acumulada no mês (% a.m.) - precisa converter
  // Usamos 4390 que já vem em % a.a.
  const val=await fetchBCBSerie(4390);
  return val;
}

async function fetchFocusIPCA(){
  // Série 13522 = IPCA acumulado 12 meses esperado (Focus)
  // Série 433   = IPCA acumulado mês (% a.m.) - precisa converter
  // Fallback: série 4175 = expectativa IPCA 12 meses
  let val=await fetchBCBSerie(13522);
  if(val==null) val=await fetchBCBSerie(4175);
  return val;
}

async function renderPatrimonio(){
  // Calcula base histórica
  const db=visibleDB();
  const cats=loadCats();
  const now=new Date();
  const y=now.getFullYear(), m=now.getMonth();

  // últimos 12 meses
  const last12Months=[];
  for(let i=11;i>=0;i--){
    let mo=m-i, yr=y;
    if(mo<0){mo+=12;yr--;}
    last12Months.push({m:mo,y:yr});
  }

  // categorias de receita excluídas/incluídas
  const allIncCats=cats.income||[];
  const includedCats=allIncCats.filter(c=>!isExcludedFromProjection(c.name));
  const excludedCats=allIncCats.filter(c=>isExcludedFromProjection(c.name));

  // receita operacional média (só categorias incluídas)
  const includedIds=new Set(includedCats.map(c=>c.id));
  let totalRec=0;
  last12Months.forEach(({m:mo,y:yr})=>{
    const txs=filterTx(db.transactions,mo,yr);
    totalRec+=txs.filter(t=>t.type==='income'&&includedIds.has(t.category)).reduce((s,t)=>s+t.amount,0);
  });
  const recBase=totalRec/12;

  // despesa média (todas)
  let totalDesp=0;
  last12Months.forEach(({m:mo,y:yr})=>{
    const txs=filterTx(db.transactions,mo,yr);
    totalDesp+=txs.filter(t=>t.type==='expense').reduce((s,t)=>s+t.amount,0);
  });
  const despBase=totalDesp/12;

  // Atualiza UI base
  document.getElementById('pat-rec-base').textContent=fmt(recBase);
  document.getElementById('pat-desp-base').textContent=fmt(despBase);
  document.getElementById('pat-rec-cats').textContent=includedCats.map(c=>c.name).join(', ')||'Nenhuma';
  document.getElementById('pat-excluded-cats').innerHTML=excludedCats.map(c=>
    `<span style="display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:20px;background:var(--bg3);border:1px solid var(--border);font-size:11px;color:var(--text3)">${c.icon} ${c.name}</span>`
  ).join('')||'<span style="font-size:11px;color:var(--text3)">Nenhuma excluída</span>';

  // Busca SELIC e IPCA
  const selicBadge=document.getElementById('pat-selic-badge');
  const ipcaBadge=document.getElementById('pat-ipca-badge');
  selicBadge.className='pat-api-badge pat-api-load'; selicBadge.textContent='⏳ Buscando SELIC...';
  ipcaBadge.className='pat-api-badge pat-api-load'; ipcaBadge.textContent='⏳ Buscando IPCA...';

  const [selicVal,ipcaVal]=await Promise.all([fetchFocusSelic(),fetchFocusIPCA()]);

  if(selicVal!=null){
    const selicInput=document.getElementById('pat-selic');
    if(!selicInput.value) selicInput.value=selicVal.toFixed(2);
    document.getElementById('pat-selic-hint').textContent=`BCB (Meta Copom): ${selicVal.toFixed(2)}% a.a.`;
    selicBadge.className='pat-api-badge pat-api-ok'; selicBadge.textContent=`✓ SELIC ${selicVal.toFixed(2)}% a.a.`;
  } else {
    selicBadge.className='pat-api-badge pat-api-err'; selicBadge.textContent='✗ API indisponível';
    document.getElementById('pat-selic-hint').textContent='Consulte bcb.gov.br e insira manualmente (ex: 13.75)';
    // Sugestão do valor atual baseado no contexto (2026)
    const selicInput=document.getElementById('pat-selic');
    if(!selicInput.value) selicInput.placeholder='Ex: 13.75';
  }

  if(ipcaVal!=null){
    const ipcaInput=document.getElementById('pat-ipca');
    if(!ipcaInput.value) ipcaInput.value=ipcaVal.toFixed(2);
    document.getElementById('pat-ipca-hint').textContent=`BCB (Focus 12m): ${ipcaVal.toFixed(2)}% a.a.`;
    ipcaBadge.className='pat-api-badge pat-api-ok'; ipcaBadge.textContent=`✓ IPCA ${ipcaVal.toFixed(2)}% a.a.`;
  } else {
    ipcaBadge.className='pat-api-badge pat-api-err'; ipcaBadge.textContent='✗ API indisponível';
    document.getElementById('pat-ipca-hint').textContent='Consulte bcb.gov.br e insira manualmente (ex: 5.50)';
    const ipcaInput=document.getElementById('pat-ipca');
    if(!ipcaInput.value) ipcaInput.placeholder='Ex: 5.50';
  }

  // guarda base para o cálculo
  _patData={recBase,despBase,includedCats,excludedCats};
}

function calcularPatrimonio(){
  if(!_patData){alert('Aguarde o carregamento dos dados.');return}
  const saldoInicial=parseFloat(document.getElementById('pat-saldo').value)||0;
  const anos=parseInt(document.getElementById('pat-anos').value)||10;
  const selicAnual=parseFloat(document.getElementById('pat-selic').value)||0;
  const ipcaAnual=parseFloat(document.getElementById('pat-ipca').value)||0;
  const crescRec=parseFloat(document.getElementById('pat-cresc-rec').value)||0;
  const crescDesp=parseFloat(document.getElementById('pat-cresc-desp').value)||0;
  const aporte=parseFloat(document.getElementById('pat-aporte').value)||0;

  if(!saldoInicial){alert('Informe o saldo atual aplicado.');return}
  if(!selicAnual){alert('Informe a taxa SELIC.');return}

  const selicMensal=(Math.pow(1+selicAnual/100,1/12)-1);
  const ipcaMensal=(Math.pow(1+ipcaAnual/100,1/12)-1);

  const {recBase,despBase}=_patData;

  // Projeção mês a mês
  let saldo=saldoInicial;
  let rendimentoAcum=0;
  let deflator=1; // para valor real

  const monthly=[];
  const yearly=[];

  for(let ano=1;ano<=anos;ano++){
    const recAnual=recBase*Math.pow(1+crescRec/100,ano-1);
    const despAnual=despBase*Math.pow(1+crescDesp/100,ano-1);
    const recMes=recAnual;
    const despMes=despAnual;

    let saldoInicioAno=saldo;
    let rendAnual=0;
    let recTotalAno=0;
    let despTotalAno=0;

    for(let mes=1;mes<=12;mes++){
      const rend=saldo*selicMensal;
      const fluxo=recMes-despMes+aporte;
      saldo=saldo+rend+fluxo;
      rendimentoAcum+=rend;
      rendAnual+=rend;
      recTotalAno+=recMes;
      despTotalAno+=despMes;
      deflator*=(1+ipcaMensal);

      monthly.push({
        ano,mes,
        saldo,
        saldobReal:saldo/deflator,
        rend,
        rec:recMes,
        desp:despMes,
        fluxo,
        rendimentoAcum
      });
    }

    yearly.push({
      ano,
      saldoInicio:saldoInicioAno,
      saldoFim:saldo,
      saldoReal:saldo/deflator,
      rendAnual,
      recTotal:recTotalAno,
      despTotal:despTotalAno,
      fluxoTotal:recTotalAno-despTotalAno+(aporte*12),
      rendimentoAcum
    });
  }

  // Marcos
  const milestones=[100000,250000,500000,1000000,2000000,5000000];
  const milestoneResults=milestones.map(target=>{
    const hit=monthly.find(m=>m.saldo>=target);
    return{target,hit};
  });

  // Atualiza cards
  const y1=yearly[0];
  const yMid=yearly[Math.min(4,yearly.length-1)];
  const yLong=yearly[Math.min(9,yearly.length-1)];
  const yFinal=yearly[yearly.length-1];

  document.getElementById('pat-cards').style.display='grid';
  document.getElementById('pat-1y').textContent=fmt(y1.saldoFim);
  document.getElementById('pat-1y-real').textContent='Real: '+fmt(y1.saldoReal);

  const midLabel=`Patrimônio em ${Math.min(5,anos)} ano${anos>=5?'s':''}`;
  document.getElementById('pat-mid-label').textContent=midLabel;
  document.getElementById('pat-mid').textContent=fmt(yMid.saldoFim);
  document.getElementById('pat-mid-real').textContent='Real: '+fmt(yMid.saldoReal);

  const longLabel=`Patrimônio em ${Math.min(10,anos)} ano${anos>=10?'s':''}`;
  document.getElementById('pat-long-label').textContent=longLabel;
  document.getElementById('pat-long').textContent=fmt(yLong.saldoFim);
  document.getElementById('pat-long-real').textContent='Real: '+fmt(yLong.saldoReal);

  document.getElementById('pat-rend-total').textContent=fmt(yFinal.rendimentoAcum);
  const pctGain=((yFinal.rendimentoAcum/saldoInicial)*100).toFixed(1);
  document.getElementById('pat-rend-pct').textContent=`+${pctGain}% sobre o capital inicial`;

  // Marcos
  const mileEl=document.getElementById('pat-milestones');
  const mileList=document.getElementById('pat-milestone-list');
  mileEl.style.display='block';
  mileList.innerHTML=milestoneResults.map(({target,hit})=>{
    const reached=!!hit;
    const label=target>=1000000?`R$ ${(target/1000000).toFixed(0)}M`:`R$ ${(target/1000).toFixed(0)}k`;
    const when=hit?`Ano ${hit.ano}, mês ${hit.mes}`:'Além do horizonte';
    return`<span class="pat-milestone ${reached?'reached':''}">
      ${reached?'✅':'⬜'} ${label}
      <span style="font-size:10px;color:${reached?'inherit':'var(--text3)'}">· ${when}</span>
    </span>`;
  }).join('');

  // Gráfico
  document.getElementById('pat-chart-card').style.display='block';
  chartPatrimonio=dc(chartPatrimonio);
  chartPatrimonio=new Chart(document.getElementById('chart-patrimonio'),{
    type:'line',
    data:{
      labels:yearly.map(y=>'Ano '+y.ano),
      datasets:[
        {
          label:'Patrimônio Nominal',
          data:yearly.map(y=>y.saldoFim),
          borderColor:'rgba(34,211,160,0.9)',
          backgroundColor:'rgba(34,211,160,0.08)',
          borderWidth:2.5,pointRadius:3,tension:0.3,fill:true
        },
        {
          label:'Valor Real (−IPCA)',
          data:yearly.map(y=>y.saldoReal),
          borderColor:'rgba(96,165,250,0.9)',
          backgroundColor:'rgba(96,165,250,0.05)',
          borderWidth:2,pointRadius:2,tension:0.3,
          borderDash:[5,3],fill:true
        },
        {
          label:'Rendimento Acumulado',
          data:yearly.map(y=>y.rendimentoAcum),
          borderColor:'rgba(251,191,36,0.8)',
          backgroundColor:'transparent',
          borderWidth:1.5,pointRadius:2,tension:0.3,
          borderDash:[3,3]
        }
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:ctx=>` ${fmt(ctx.raw)}`}}
      },
      scales:{
        x:{grid:{color:'rgba(255,255,255,0.04)'},ticks:{color:'#8b93a8',maxTicksLimit:12}},
        y:{grid:{color:'rgba(255,255,255,0.04)'},ticks:{color:'#8b93a8',callback:v=>{
          if(v>=1000000)return'R$'+(v/1000000).toFixed(1)+'M';
          if(v>=1000)return'R$'+(v/1000).toFixed(0)+'k';
          return'R$'+v;
        }}}
      }
    }
  });

  // Tabela — modo anual por padrão
  document.getElementById('pat-table-card').style.display='block';
  document.getElementById('pat-show-monthly').checked=false;

  // guarda dados para toggle mensal
  window._patYearly=yearly;
  window._patMonthly=monthly;

  renderPatTable(false);
}

function renderPatTable(showMonthly){
  const tbody=document.getElementById('pat-tbody');
  if(!tbody)return;

  if(!showMonthly){
    tbody.innerHTML=(window._patYearly||[]).map((y,i)=>{
      const isHighlight=(y.ano===1||y.ano===5||y.ano===10);
      const balColor=y.fluxoTotal>=0?'var(--green)':'var(--red)';
      return`<tr class="${isHighlight?'highlight-year':''}">
        <td>Ano ${y.ano}</td>
        <td style="color:var(--green)">${fmt(y.recTotal)}</td>
        <td style="color:var(--red)">${fmt(y.despTotal)}</td>
        <td style="color:${balColor}">${y.fluxoTotal>=0?'+':''}${fmt(y.fluxoTotal)}</td>
        <td style="color:var(--amber)">${fmt(y.rendAnual)}</td>
        <td style="color:var(--green);font-weight:600">${fmt(y.saldoFim)}</td>
        <td style="color:var(--blue)">${fmt(y.saldoReal)}</td>
      </tr>`;
    }).join('');
  } else {
    tbody.innerHTML=(window._patMonthly||[]).map(m=>{
      const balColor=m.fluxo>=0?'var(--green)':'var(--red)';
      return`<tr>
        <td style="color:var(--text2);font-weight:400">Ano ${m.ano} · ${MONTHS_SHORT[m.mes-1]}</td>
        <td style="color:var(--green)">${fmt(m.rec)}</td>
        <td style="color:var(--red)">${fmt(m.desp)}</td>
        <td style="color:${balColor}">${m.fluxo>=0?'+':''}${fmt(m.fluxo)}</td>
        <td style="color:var(--amber)">${fmt(m.rend)}</td>
        <td style="color:var(--green);font-weight:600">${fmt(m.saldo)}</td>
        <td style="color:var(--blue)">${fmt(m.saldobReal)}</td>
      </tr>`;
    }).join('');
  }
}

function toggleMonthlyDetail(){
  const show=document.getElementById('pat-show-monthly')?.checked;
  renderPatTable(show);
}

// ===================== INIT =====================
(async function init(){
  await loadAll();
  populateAllSelects();
  if(_fbUrl){const b=document.getElementById('sync-status-badge');if(b){b.className='sync-status sync-ok';b.innerHTML='<span class="sync-dot ok"></span>Firebase conectado'}}
  const last=localStorage.getItem(DEVICE_PROFILE_KEY);
  if(last&&PROFILES[last])pickProfile(last);
})();
</script>
</body>
</html>
