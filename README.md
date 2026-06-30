<!DOCTYPE html>
<html lang="en" data-theme="light">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Srinivasa Heart Centre | Clinic Portal</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<!-- SheetJS for Excel Export -->
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<!-- EmailJS SDK -->
<script src="https://cdn.jsdelivr.net/npm/@emailjs/browser@4/dist/email.min.js"></script>
<!-- QRCode.js for UPI payment QR codes -->
<script src="https://cdn.jsdelivr.net/npm/qrcode@1.5.3/build/qrcode.min.js"></script>
<style>
/* ===== ANIMATED MOVING BACKGROUND ===== */
@keyframes bgMove {
  0%   { background-position: 0% 50%; }
  50%  { background-position: 100% 50%; }
  100% { background-position: 0% 50%; }
}
@keyframes floatBubble {
  0%   { transform: translateY(0px) scale(1); opacity: 0.15; }
  50%  { transform: translateY(-30px) scale(1.1); opacity: 0.25; }
  100% { transform: translateY(0px) scale(1); opacity: 0.15; }
}
.animated-bg {
  position: fixed;
  top: 0; left: 0;
  width: 100vw; height: 100vh;
  z-index: -1;
  background: linear-gradient(270deg, #00f2ff, #7628d0, #a63636, #1a2163, #cf2c88, #0d0f7e);
  background-size: 400% 400%;
  animation: bgMove 12s ease infinite;
  pointer-events: none;
}
[data-theme="dark"] .animated-bg {
  background: linear-gradient(270deg, #081129, #080c21, #060835, #0e1b38, #0e1a30, #091525);
  background-size: 400% 400%;
  animation: bgMove 12s ease infinite;
}
.bubble {
  position: fixed;
  border-radius: 50%;
  pointer-events: none;
  z-index: -1;
  filter: blur(40px);
}
.bubble-1 {
  width: 400px; height: 400px;
  background: radial-gradient(circle, rgba(6,182,212,0.3), transparent);
  top: -100px; left: -100px;
  animation: floatBubble 8s ease-in-out infinite;
}
.bubble-2 {
  width: 300px; height: 300px;
  background: radial-gradient(circle, rgba(139,92,246,0.2), transparent);
  top: 30%; right: -80px;
  animation: floatBubble 10s ease-in-out infinite 2s;
}
.bubble-3 {
  width: 350px; height: 350px;
  background: radial-gradient(circle, rgba(16,185,129,0.2), transparent);
  bottom: -80px; left: 30%;
  animation: floatBubble 9s ease-in-out infinite 4s;
}
.bubble-4 {
  width: 250px; height: 250px;
  background: radial-gradient(circle, rgba(239,68,68,0.15), transparent);
  bottom: 20%; right: 20%;
  animation: floatBubble 11s ease-in-out infinite 1s;
}
 
/* ===== CSS VARIABLES ===== */
:root {
  --primary: #06b6d4;
  --primary-hover: #0891b2;
  --primary-dark: #0e7490;
  --primary-light: #ecfeff;
  --secondary: #0f172a;
  --secondary-light: #1e293b;
  --accent: #ef4444;
  --accent-hover: #e7d5d5;
  --accent-light: #0400f3;
  --success: #e30000;
  --success-light: #d1fae5;
  --warning: #f59e0b;
  --warning-light: #fef3c7;
  --bg-page: rgba(248,250,252,0.85);
  --bg-card: rgba(255,255,255,0.92);
  --border-color: #e2e8f0;
  --text-primary: #0f172a;
  --text-secondary: #475569;
  --text-muted: #94a3b8;
  --shadow-sm: 0 1px 2px 0 rgba(0,0,0,0.05);
  --shadow-md: 0 4px 6px -1px rgba(0,0,0,0.07),0 2px 4px -2px rgba(0,0,0,0.05);
  --shadow-lg: 0 10px 15px -3px rgba(0,0,0,0.1),0 4px 6px -4px rgba(0,0,0,0.1);
  --radius-sm: 8px;
  --radius-md: 12px;
  --radius-lg: 20px;
  --transition: all 0.25s cubic-bezier(0.4,0,0.2,1);
}
[data-theme="dark"] {
  --primary: #22d3ee;
  --primary-hover: #06b6d4;
  --primary-dark: #0891b2;
  --primary-light: #083344;
  --accent: #f87171;
  --accent-hover: #ef4444;
  --accent-light: #450a0a;
  --success: #34d399;
  --success-light: #064e3b;
  --warning: #fbbf24;
  --warning-light: #78350f;
  --bg-page: rgba(11,19,41,0.88);
  --bg-card: rgba(17,28,53,0.92);
  --border-color: #1e2d4a;
  --text-primary: #f8fafc;
  --text-secondary: #94a3b8;
  --text-muted: #64748b;
  --shadow-sm: 0 1px 2px 0 rgba(0,0,0,0.2);
  --shadow-md: 0 4px 6px -1px rgba(0,0,0,0.3),0 2px 4px -2px rgba(0,0,0,0.3);
  --shadow-lg: 0 10px 15px -3px rgba(0,0,0,0.4),0 4px 6px -4px rgba(0,0,0,0.4);
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Plus Jakarta Sans',sans-serif;background-color:transparent;color:var(--text-primary);min-height:100vh;display:flex;transition:var(--transition);overflow:hidden;}
 
/* SIDEBAR */
.sidebar{width:260px;background-color:rgba(11,19,41,0.97);backdrop-filter:blur(10px);color:#f8fafc;height:100vh;position:fixed;left:0;top:0;display:flex;flex-direction:column;z-index:100;box-shadow:4px 0 20px rgba(0,0,0,0.2);transition:var(--transition);}
.sidebar-brand{padding:24px 20px;display:flex;align-items:center;gap:12px;border-bottom:1px solid rgba(255,255,255,0.05);}
.sidebar-logo{width:40px;height:40px;background:linear-gradient(135deg,var(--primary),#0284c7);border-radius:10px;display:flex;align-items:center;justify-content:center;font-weight:700;color:white;font-size:20px;box-shadow:0 4px 12px rgba(6,182,212,0.3);position:relative;}
.sidebar-logo::after{content:'';position:absolute;width:10px;height:10px;background:#996d1a;border-radius:50%;bottom:-2px;right:-2px;border:2px solid #0b1329;animation:heartPulse 1.5s infinite;}
@keyframes heartPulse{0%{transform:scale(1);opacity:1;}50%{transform:scale(1.3);opacity:0.5;}100%{transform:scale(1);opacity:1;}}
.brand-details h2{font-family:'Outfit',sans-serif;font-size:14px;font-weight:700;letter-spacing:0.5px;line-height:1.2;}
.brand-details p{font-size:10px;color:var(--primary);font-weight:600;text-transform:uppercase;letter-spacing:0.8px;margin-top:2px;}
.sidebar-menu{list-style:none;padding:20px 12px;display:flex;flex-direction:column;gap:6px;overflow-y:auto;flex:1;}
.sidebar-item a{display:flex;align-items:center;gap:12px;padding:12px 16px;color:#6c82a0;text-decoration:none;font-size:13px;font-weight:600;border-radius:8px;transition:var(--transition);}
.sidebar-item a svg{width:18px;height:18px;transition:var(--transition);}
.sidebar-item a:hover{color:#ffffff;background-color:rgba(255,255,255,0.04);}
.sidebar-item.active a{color:#ffffff;background:linear-gradient(135deg,rgba(6,182,212,0.15),rgba(6,182,212,0.05));border-left:3px solid var(--primary);padding-left:13px;}
.sidebar-item.active a svg{color:var(--primary);}
.sidebar-footer{padding:15px 20px;border-top:1px solid rgba(255,255,255,0.05);background-color:rgba(0,0,0,0.15);font-size:11px;color:#64748b;display:flex;flex-direction:column;gap:4px;}
.sidebar-footer strong{color:#cbd5e1;}
 
/* CONTENT */
.content{margin-left:260px;flex:1;height:100vh;display:flex;flex-direction:column;overflow-y:auto;position:relative;transition:var(--transition);}
.top-header{background-color:var(--bg-card);backdrop-filter:blur(10px);border-bottom:1px solid var(--border-color);padding:15px 30px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:90;transition:var(--transition);}
.header-meta h1{font-family:'Outfit',sans-serif;font-size:18px;font-weight:700;color:var(--text-primary);}
.header-meta p{font-size:12px;color:var(--text-secondary);}
.header-actions{display:flex;align-items:center;gap:15px;}
.theme-toggle-btn{background-color:var(--bg-page);border:1px solid var(--border-color);color:var(--text-primary);width:38px;height:38px;border-radius:8px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:var(--transition);}
.theme-toggle-btn:hover{border-color:var(--primary);background-color:var(--primary-light);color:var(--primary);}
.dr-badge{display:flex;align-items:center;gap:10px;padding:6px 14px;background-color:var(--primary-light);border:1px solid rgba(6,182,212,0.15);border-radius:30px;}
.dr-avatar{width:26px;height:26px;background:linear-gradient(135deg,var(--primary),var(--primary-hover));color:white;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;}
.dr-name{font-size:12px;font-weight:700;color:var(--text-primary);}
.page-container{padding:25px 30px;flex:1;overflow-y:auto;display:flex;flex-direction:column;gap:25px;}
.page-view{display:none;animation:pageFadeIn 0.3s cubic-bezier(0.16,1,0.3,1) forwards;}
.page-view.active-view{display:flex;flex-direction:column;gap:25px;}
@keyframes pageFadeIn{from{opacity:0;transform:translateY(8px);}to{opacity:1;transform:translateY(0);}}
 
/* CARDS */
.card{background-color:var(--bg-card);backdrop-filter:blur(8px);border:1px solid var(--border-color);border-radius:var(--radius-md);box-shadow:var(--shadow-sm);padding:24px;display:flex;flex-direction:column;gap:20px;transition:var(--transition);}
.card:hover{box-shadow:var(--shadow-md);}
.card-header{display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--border-color);padding-bottom:15px;}
.card-title{font-family:'Outfit',sans-serif;font-size:16px;font-weight:700;color:var(--text-primary);display:flex;align-items:center;gap:10px;}
.card-title svg{color:var(--primary);width:20px;height:20px;}
 
/* STATS */
.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:20px;}
.stat-card{background-color:var(--bg-card);backdrop-filter:blur(8px);border:1px solid var(--border-color);border-radius:var(--radius-md);padding:20px;display:flex;justify-content:space-between;align-items:center;box-shadow:var(--shadow-sm);transition:var(--transition);}
.stat-card:hover{transform:translateY(-2px);box-shadow:var(--shadow-md);border-color:var(--primary);}
.stat-info h3{font-size:12px;color:var(--text-secondary);font-weight:600;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;}
.stat-value{font-family:'Outfit',sans-serif;font-size:28px;font-weight:700;color:var(--text-primary);}
.stat-change{font-size:11px;font-weight:600;display:flex;align-items:center;gap:4px;margin-top:4px;}
.stat-change.up{color:var(--success);}
.stat-change.down{color:var(--accent);}
.stat-icon{width:48px;height:48px;border-radius:10px;display:flex;align-items:center;justify-content:center;}
.stat-card:nth-child(1) .stat-icon{background-color:rgba(6,182,212,0.1);color:var(--primary);}
.stat-card:nth-child(2) .stat-icon{background-color:rgba(16,185,129,0.1);color:var(--success);}
.stat-card:nth-child(3) .stat-icon{background-color:rgba(245,158,11,0.1);color:var(--warning);}
.stat-card:nth-child(4) .stat-icon{background-color:rgba(239,68,68,0.1);color:var(--accent);}
.stat-icon svg{width:24px;height:24px;}
.grid-2col{display:grid;grid-template-columns:1fr 1fr;gap:25px;}
@media(max-width:992px){.grid-2col{grid-template-columns:1fr;}}
 
/* FORMS */
.form-section{border:1px solid var(--border-color);padding:20px;border-radius:var(--radius-md);margin-bottom:20px;background-color:rgba(248,250,252,0.4);transition:var(--transition);}
[data-theme="dark"] .form-section{background-color:rgba(9,13,22,0.2);}
.form-section-title-bar{border-left:4px solid var(--primary);padding-left:10px;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center;}
.form-section h3{font-family:'Outfit',sans-serif;font-size:15px;font-weight:700;color:var(--text-primary);text-transform:uppercase;letter-spacing:0.5px;}
.form-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:15px;margin-bottom:15px;}
.form-col-3{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;margin-bottom:15px;}
.form-group{display:flex;flex-direction:column;gap:6px;margin-bottom:12px;}
.form-group.full-width{grid-column:span 100%;}
label{font-size:11px;font-weight:700;color:var(--text-secondary);text-transform:uppercase;letter-spacing:0.5px;}
input[type="text"],input[type="number"],input[type="email"],input[type="date"],select,textarea{font-family:inherit;font-size:13px;padding:10px 14px;color:var(--text-primary);background-color:var(--bg-page);border:1px solid var(--border-color);border-radius:var(--radius-sm);transition:var(--transition);outline:none;width:100%;}
input:focus,select:focus,textarea:focus{border-color:var(--primary);background-color:var(--bg-card);box-shadow:0 0 0 3px rgba(6,182,212,0.15);}
input::placeholder,textarea::placeholder{color:var(--text-muted);}
textarea{resize:vertical;min-height:70px;}
.checkbox-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:10px;margin:8px 0 15px 0;}
.checkbox-label{display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13px;font-weight:600;color:var(--text-primary);background-color:var(--bg-page);border:1px solid var(--border-color);padding:10px 12px;border-radius:var(--radius-sm);transition:var(--transition);text-transform:none;letter-spacing:normal;}
.checkbox-label input[type="checkbox"]{width:16px;height:16px;accent-color:var(--primary);cursor:pointer;}
.checkbox-label:hover{border-color:var(--primary);background-color:var(--primary-light);}
.checkbox-label.checked{border-color:var(--primary);background-color:var(--primary-light);color:var(--primary-dark);}
 
/* WIZARD */
.wizard-header{display:flex;justify-content:space-between;align-items:center;position:relative;margin-bottom:30px;padding:0 10px;}
.wizard-header::before{content:'';position:absolute;height:3px;background-color:var(--border-color);top:50%;left:40px;right:40px;transform:translateY(-50%);z-index:1;}
.wizard-line-progress{position:absolute;height:3px;background-color:var(--primary);top:50%;left:40px;transform:translateY(-50%);z-index:2;width:0%;transition:width 0.3s ease;}
.wizard-step-node{display:flex;flex-direction:column;align-items:center;gap:8px;z-index:3;position:relative;background-color:var(--bg-card);padding:0 10px;cursor:pointer;}
.wizard-node-circle{width:36px;height:36px;border-radius:50%;background-color:var(--bg-page);border:2px solid var(--border-color);color:var(--text-muted);font-size:13px;font-weight:700;display:flex;align-items:center;justify-content:center;transition:var(--transition);}
.wizard-step-node.active-node .wizard-node-circle{border-color:var(--primary);background-color:var(--primary);color:white;box-shadow:0 0 0 4px rgba(6,182,212,0.15);}
.wizard-step-node.completed-node .wizard-node-circle{border-color:var(--success);background-color:var(--success);color:white;}
.wizard-node-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:var(--text-muted);transition:var(--transition);}
.wizard-step-node.active-node .wizard-node-label{color:var(--primary);}
.wizard-step-node.completed-node .wizard-node-label{color:var(--success);}
 
/* FILE UPLOAD */
.file-upload-area{border:2px dashed var(--border-color);background-color:var(--bg-page);border-radius:var(--radius-sm);padding:20px;text-align:center;cursor:pointer;transition:var(--transition);position:relative;}
.file-upload-area:hover{border-color:var(--primary);background-color:var(--primary-light);}
.file-upload-area input[type="file"]{position:absolute;width:100%;height:100%;top:0;left:0;opacity:0;cursor:pointer;}
.file-upload-icon{font-size:28px;margin-bottom:8px;color:var(--text-muted);}
.file-upload-text{font-size:12px;font-weight:600;color:var(--text-secondary);}
.file-upload-text span{color:var(--primary);text-decoration:underline;}
.file-preview-list{margin-top:10px;display:flex;flex-direction:column;gap:8px;}
.file-preview-item{display:flex;align-items:center;justify-content:space-between;padding:8px 12px;background-color:var(--bg-page);border:1px solid var(--border-color);border-radius:var(--radius-sm);font-size:12px;}
.file-preview-details{display:flex;align-items:center;gap:8px;}
.file-preview-icon{color:var(--primary);}
.file-preview-name{font-weight:600;}
.file-preview-remove{background:transparent;border:none;color:var(--accent);cursor:pointer;font-weight:700;}
.photo-uploader{display:flex;align-items:center;gap:15px;margin-bottom:15px;}
.photo-preview{width:80px;height:80px;border-radius:50%;border:2px solid var(--border-color);background-color:var(--bg_page);background-size:cover;background-position:center;display:flex;align-items:center;justify-content:center;font-size:24px;color:var(--text-muted);overflow:hidden;}
.photo-input-wrapper{flex:1;}
 
/* BUTTONS */
.btn{font-family:inherit;font-size:13px;font-weight:700;padding:10px 18px;border-radius:var(--radius-sm);cursor:pointer;border:none;display:inline-flex;align-items:center;justify-content:center;gap:8px;transition:var(--transition);}
.btn-primary{background-color:var(--primary);color:white;box-shadow:0 4px 12px rgba(6,182,212,0.15);}
.btn-primary:hover{background-color:var(--primary-hover);transform:translateY(-1px);}
.btn-secondary{background-color:var(--bg-page);border:1px solid var(--border-color);color:var(--text-primary);}
.btn-secondary:hover{background-color:var(--border-color);}
.btn-accent{background-color:var(--accent);color:white;}
.btn-accent:hover{background-color:var(--accent-hover);}
.btn-success{background-color:var(--success);color:white;}
.btn-success:hover{background-color:#059669;}
.btn-icon{padding:8px;border-radius:6px;display:flex;align-items:center;justify-content:center;}
 
/* Excel Export Button Special Style */
.btn-excel{background:linear-gradient(135deg,#ad2a2a,#240078);color:white;box-shadow:0 4px 12px rgba(33,163,102,0.3);}
.btn-excel:hover{background:linear-gradient(135deg,#165c35,#1d8f58);transform:translateY(-1px);}
 
/* TABLE */
.table-controls{display:flex;justify-content:space-between;align-items:center;gap:15px;flex-wrap:wrap;}
.search-bar-wrapper{position:relative;max-width:360px;width:100%;}
.search-bar-wrapper input{padding-left:38px;}
.search-bar-wrapper svg{position:absolute;left:12px;top:50%;transform:translateY(-50%);width:18px;height:18px;color:var(--text-muted);pointer-events:none;}
.filters-wrapper{display:flex;gap:8px;flex-wrap:wrap;}
.filter-chip{padding:6px 14px;border-radius:30px;border:1px solid var(--border-color);background-color:var(--bg-card);font-size:11px;font-weight:700;cursor:pointer;transition:var(--transition);color:var(--text-secondary);}
.filter-chip:hover{border-color:var(--primary);color:var(--primary);}
.filter-chip.active-chip{background-color:var(--primary);border-color:var(--primary);color:white;}
.table-wrapper{width:100%;overflow-x:auto;border:1px solid var(--border-color);border-radius:var(--radius-md);background-color:var(--bg-card);}
table{width:100%;border-collapse:collapse;text-align:left;font-size:13px;}
th{background-color:var(--bg-page);color:var(--text-secondary);font-weight:700;padding:14px 18px;border-bottom:1px solid var(--border-color);text-transform:uppercase;font-size:10px;letter-spacing:0.5px;}
td{padding:14px 18px;border-bottom:1px solid var(--border-color);color:var(--text-primary);vertical-align:middle;}
tr:last-child td{border-bottom:none;}
tr{transition:var(--transition);}
tr:hover{background-color:rgba(6,182,212,0.02);}
 
/* BADGES */
.badge{padding:4px 10px;border-radius:20px;font-size:10px;font-weight:700;text-transform:uppercase;display:inline-flex;align-items:center;gap:4px;width:fit-content;}
.badge-blue{background-color:var(--primary-light);color:var(--primary-dark);}
.badge-green{background-color:var(--success-light);color:var(--success);}
.badge-red{background-color:var(--accent-light);color:var(--accent);}
.badge-yellow{background-color:var(--warning-light);color:var(--warning);}
 
/* MODAL */
.modal-overlay{position:fixed;top:0;left:0;width:100vw;height:100vh;background-color:rgba(9,13,22,0.6);backdrop-filter:blur(4px);display:none;align-items:center;justify-content:center;z-index:1000;padding:20px;}
.modal-content{background-color:var(--bg-card);border:1px solid var(--border-color);width:100%;max-width:900px;max-height:90vh;border-radius:var(--radius-md);box-shadow:var(--shadow-lg);display:flex;flex-direction:column;overflow:hidden;animation:modalContentSlide 0.3s cubic-bezier(0.16,1,0.3,1) forwards;}
@keyframes modalContentSlide{from{transform:translateY(20px);opacity:0;}to{transform:translateY(0);opacity:1;}}
.modal-header{padding:20px 24px;border-bottom:1px solid var(--border-color);display:flex;justify-content:space-between;align-items:center;}
.modal-body{padding:24px;overflow-y:auto;display:flex;flex-direction:column;gap:25px;}
.modal-footer{padding:15px 24px;border-top:1px solid var(--border-color);display:flex;justify-content:flex-end;gap:10px;background-color:var(--bg-page);}
 
/* EDIT MODAL - wider */
.edit-modal-content{max-width:1000px;}
 
/* CLINICAL */
.clinical-header{display:flex;align-items:center;gap:20px;border-bottom:1px dashed var(--border-color);padding-bottom:20px;}
.clinical-avatar{width:70px;height:70px;border-radius:50%;background-color:var(--primary-light);display:flex;align-items:center;justify-content:center;font-size:24px;border:2px solid var(--primary);overflow:hidden;background-size:cover;background-position:center;}
.clinical-meta-info{flex:1;}
.clinical-meta-info h2{font-family:'Outfit',sans-serif;font-size:20px;font-weight:700;}
.clinical-meta-details{display:flex;gap:15px;font-size:12px;color:var(--text-secondary);margin-top:5px;flex-wrap:wrap;}
.clinical-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;}
@media(max-width:600px){.clinical-grid{grid-template-columns:1fr;}}
.clinical-item{background-color:var(--bg-page);border:1px solid var(--border-color);padding:12px;border-radius:var(--radius-sm);}
.clinical-label{font-size:10px;font-weight:700;text-transform:uppercase;color:var(--text-muted);letter-spacing:0.5px;}
.clinical-value{font-size:13px;font-weight:600;color:var(--text-primary);margin-top:3px;}
 
/* SLOTS */
.scheduler-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(100px,1fr));gap:10px;margin-top:8px;}
.slot-btn{border:1px solid var(--border-color);background-color:var(--bg-card);padding:10px;border-radius:var(--radius-sm);text-align:center;font-size:12px;font-weight:700;cursor:pointer;transition:var(--transition);}
.slot-btn:hover{border-color:var(--primary);background-color:var(--primary-light);}
.slot-btn.active-slot{background-color:var(--primary);border-color:var(--primary);color:white;}
.slot-btn.disabled-slot{background-color:var(--bg-page);color:var(--text-muted);cursor:not-allowed;border-style:dashed;}
 
/* APPOINTMENTS */
.planner-board{display:flex;flex-direction:column;gap:15px;}
.appointment-card{background-color:var(--bg-card);border:1px solid var(--border-color);border-left:5px solid var(--primary);padding:16px;border-radius:var(--radius-md);display:flex;justify-content:space-between;align-items:center;transition:var(--transition);}
.appointment-card.urgent{border-left-color:var(--accent);}
.appointment-card:hover{transform:translateX(4px);}
.ap-time{font-family:'Outfit',sans-serif;font-size:14px;font-weight:700;color:var(--primary);min-width:80px;}
.ap-details h4{font-size:14px;font-weight:700;}
.ap-details p{font-size:12px;color:var(--text-secondary);margin-top:3px;}
.ap-actions{display:flex;gap:8px;}
 
/* TOAST */
.toast-container{position:fixed;bottom:30px;right:30px;display:flex;flex-direction:column;gap:10px;z-index:2000;}
.toast{background-color:var(--bg-card);border:1px solid var(--border-color);border-left:4px solid var(--primary);box-shadow:var(--shadow-lg);padding:16px 20px;border-radius:var(--radius-sm);font-size:13px;font-weight:600;display:flex;align-items:center;gap:10px;animation:toastSlide 0.3s cubic-bezier(0.16,1,0.3,1) forwards;min-width:300px;}
@keyframes toastSlide{from{transform:translateX(120%);}to{transform:translateX(0);}}
.toast.success{border-left-color:var(--success);}
.toast.error{border-left-color:var(--accent);}
.toast.info{border-left-color:var(--primary);}
 
/* EMAIL */
.email-container{display:grid;grid-template-columns:1.2fr 1fr;gap:25px;}
@media(max-width:992px){.email-container{grid-template-columns:1fr;}}
.email-preview-frame{border:1px solid var(--border-color);border-radius:var(--radius-md);background-color:rgb(59, 59, 59);color:#ffffff;padding:30px;box-shadow:var(--shadow-sm);display:flex;flex-direction:column;gap:15px;font-family:Arial,sans-serif;min-height:400px;}
.email-preview-header{border-bottom:2px solid #ff0000;padding-bottom:15px;display:flex;align-items:center;gap:15px;}
.email-preview-body{font-size:14px;line-height:1.6;flex:1;}
.email-preview-footer{border-top:1px solid #ff0000;padding-top:15px;font-size:11px;color:#ffffff;text-align:center;}
 
/* CALCULATORS */
.calc-card{background-color:var(--bg-card);border:1px solid var(--border-color);border-radius:var(--radius-md);padding:20px;}
.calc-range-slider{width:100%;margin:10px 0;accent-color:var(--primary);}
.calc-result-box{background-color:var(--primary-light);border:1px solid rgba(6,182,212,0.2);padding:15px;border-radius:var(--radius-sm);margin-top:15px;text-align:center;}
.calc-result-value{font-family:'Outfit',sans-serif;font-size:24px;font-weight:700;color:var(--primary-dark);}
.invoice-builder-items{display:flex;flex-direction:column;gap:8px;}
.chart-container{position:relative;height:250px;width:100%;}
 
/* EXCEL BANNER */
.excel-banner{background:linear-gradient(135deg,rgba(29,111,66,0.1),rgba(33,163,102,0.08));border:1px solid rgba(33,163,102,0.2);border-radius:var(--radius-sm);padding:12px 16px;display:flex;align-items:center;justify-content:space-between;gap:10px;margin-top:10px;}
.excel-banner-text{font-size:12px;font-weight:600;color:var(--text-secondary);}
.excel-banner-text strong{color:#1d6f42;}
[data-theme="dark"] .excel-banner-text strong{color:#34d399;}
</style>
</head>
<body>
 
<!-- ANIMATED BACKGROUND -->
<div class="animated-bg"></div>
<div class="bubble bubble-1"></div>
<div class="bubble bubble-2"></div>
<div class="bubble bubble-3"></div>       
<div class="bubble bubble-4"></div>
 
<!-- SIDEBAR -->
<div class="sidebar">
  <div class="sidebar-brand">
    <div class="sidebar-logo">❤️</div>
    <div class="brand-details">
      <h2>Srinivasa Heart</h2>
      <p>Dr. Srinivas Ramaka</p>
    </div>
  </div>
  <ul class="sidebar-menu">
    <li class="sidebar-item active" id="menu-home">
      <a href="#" onclick="switchView('home')">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="9"/><rect x="14" y="3" width="7" height="5"/><rect x="14" y="12" width="7" height="9"/><rect x="3" y="16" width="7" height="5"/></svg>
        HOME
      </a>
    </li>
    <li class="sidebar-item" id="menu-new-patient">
      <a href="#" onclick="switchView('new-patient')">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><line x1="19" y1="8" x2="19" y2="14"/><line x1="22" y1="11" x2="16" y2="11"/></svg>
        NEW PATIENT
      </a>
    </li>
    <li class="sidebar-item" id="menu-old-patient">
      <a href="#" onclick="switchView('old-patient')">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>
        OLD PATIENT
      </a>
    </li>
    <li class="sidebar-item" id="menu-appointment">
      <a href="#" onclick="switchView('appointment')">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
        BOOK APPOINTMENT
      </a>
    </li>
    <li class="sidebar-item" id="menu-appointment-view">
      <a href="#" onclick="switchView('appointment-view')">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 8v4l3 3"/><circle cx="12" cy="12" r="9"/></svg>
        APPOINTMENT VIEW
      </a>
    </li>
    <li class="sidebar-item" id="menu-study">
      <a href="#" onclick="switchView('study')">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M4 4.5A2.5 2.5 0 0 1 6.5 2H20v20H6.5a2.5 2.5 0 0 1-2.5-2.5v-15z"/></svg>
        STUDY / CLINICAL
      </a>
    </li>
    <li class="sidebar-item" id="menu-email">
      <a href="#" onclick="switchView('email')">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>
        EMAIL PORTAL
      </a>
    </li>
    <li class="sidebar-item" id="menu-accounts">
      <a href="#" onclick="switchView('accounts')">
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>
        ACCOUNTS
      </a>
    </li>
  </ul>
  <div class="sidebar-footer">
    <div>Logged in as:</div>
    <strong>Dr. Srinivas Ramaka</strong>
    <div style="font-size:9px;opacity:0.7;margin-top:5px;">MD., DM. (Cardiology)</div>
  </div>
</div>
 
<!-- MAIN CONTENT -->
<div class="content">
  <div class="top-header">
    <div class="header-meta">
      <h1 id="page-title-display">Srinivasa Heart Centre</h1>
      <p id="page-subtitle-display">Consultant Cardiologist Diagnostic Dashboard</p>
    </div>
    <div class="header-actions">
      <div style="text-align:right;margin-right:15px;font-size:11px;opacity:0.8;" id="clock-display">Loading...</div>
      <button class="theme-toggle-btn" onclick="toggleTheme()" title="Toggle Dark/Light Mode">
        <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/></svg>
      </button>
      <button class="btn btn-excel" onclick="exportAllPatientsToExcel()" title="Export All Patients to Excel" style="padding:8px 14px;font-size:12px;">
        📊 Export Excel
      </button>
      <div class="dr-badge">
        <div class="dr-avatar">SR</div>
        <div class="dr-name">Dr. Srinivas Ramaka</div>
      </div>
    </div>
  </div>
 
  <div class="page-container">
 
    <!-- HOME -->
    <div id="view-home" class="page-view active-view">
      <div class="stats-grid">
        <div class="stat-card">
          <div class="stat-info"><h3>Total Patients</h3><div class="stat-value" id="stat-total-patients">0</div><div class="stat-change up">▲ +12% this month</div></div>
          <div class="stat-icon"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/><path d="M21 21v-2a4 4 0 0 0-3-3.87"/><path d="M23 3.13a4 4 0 0 1 0 7.75"/></svg></div>
        </div>
        <div class="stat-card">
          <div class="stat-info"><h3>Today's Appts</h3><div class="stat-value" id="stat-today-appointments">0</div><div class="stat-change up">▲ +4 today</div></div>
          <div class="stat-icon"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg></div>
        </div>
        <div class="stat-card">
          <div class="stat-info"><h3>Collected Revenue</h3><div class="stat-value" id="stat-collected-revenue">₹0</div><div class="stat-change up">▲ 92% paid status</div></div>
          <div class="stat-icon"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg></div>
        </div>
      </div>
      <div class="grid-2col">
        <div class="card">
          <div class="card-header">
            <div class="card-title"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>Cardiovascular Visit Trends</div>
          </div>
          <div class="chart-container"><canvas id="dashboardChart"></canvas></div>
        </div>
        <div class="card">
          <div class="card-header">
            <div class="card-title"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>Recent Clinical Activity</div>
          </div>
          <div style="display:flex;flex-direction:column;gap:15px;" id="dashboard-recent-activity"></div>
        </div>
      </div>
    </div>
 
    <!-- NEW PATIENT -->
    <div id="view-new-patient" class="page-view">
      <div class="card">
        <div class="card-header">
          <div class="card-title"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></svg>Cardiology Patient Registry Wizard</div>
          <span class="badge badge-blue" id="wizard-step-indicator">STEP 1 OF 3</span>
        </div>
        <div class="wizard-header">
          <div class="wizard-line-progress" id="wizard-line"></div>
          <div class="wizard-step-node active-node" onclick="goToStep(1)"><div class="wizard-node-circle">1</div><div class="wizard-node-label">Personal</div></div>
          <div class="wizard-step-node" onclick="goToStep(2)"><div class="wizard-node-circle">2</div><div class="wizard-node-label">Medical</div></div>
          <div class="wizard-step-node" onclick="goToStep(3)"><div class="wizard-node-circle">3</div><div class="wizard-node-label">Exam</div></div>
        </div>
        <form id="wizard-patient-form">
          <!-- STEP 1 -->
          <div id="wizard-step-1" class="wizard-step-content" style="display:block;">
            <div class="form-section">
              <div class="form-section-title-bar"><h3>Personal Details & Demographics</h3></div>
              <div class="photo-uploader">
                <div class="photo-preview" id="photo-upload-preview">📸</div>
                <div class="photo-input-wrapper"><label>Patient Photo</label><input type="file" id="reg-photo" accept="image/*" onchange="previewPatientPhoto(event)"></div>
              </div>
              <div class="form-row">
                <div class="form-group"><label>Identity Card Type</label><select id="reg-card-type"><option value="Aadhar">Aadhar</option><option value="PAN">PAN</option><option value="Passport">Passport</option></select></div>
                <div class="form-group"><label>Card Number *</label><input type="text" id="reg-card-number" placeholder="Enter card details" required></div>
              </div>
              <div class="form-row">
                <div class="form-group"><label>First Name *</label><input type="text" id="reg-first-name" placeholder="First Name" required></div>
                <div class="form-group"><label>Last Name *</label><input type="text" id="reg-last-name" placeholder="Last Name" required></div>
              </div>
              <div class="form-row">
                <div class="form-group"><label>Date Of Birth *</label><input type="date" id="reg-dob" onchange="calculateAgeFromDOB()" required></div>
                <div class="form-group"><label>Age (Auto-calculated)</label><input type="number" id="reg-age" placeholder="Age" min="0" max="120" required></div>
              </div>
              <div class="form-row">
                <div class="form-group"><label>Gender</label><select id="reg-gender"><option value="Male">Male</option><option value="Female">Female</option><option value="Other">Other</option></select></div>
                <div class="form-group"><label>Occupation</label><input type="text" id="reg-occupation" placeholder="e.g. Engineer, Teacher"></div>
              </div>
              <div class="form-group full-width"><label>D.No / Area / Street</label><textarea id="reg-address" placeholder="Address lines"></textarea></div>
              <div class="form-row">
                <div class="form-group"><label>Home Town</label><input type="text" id="reg-hometown" placeholder="Home Town"></div>
                <div class="form-group"><label>Mandal</label><input type="text" id="reg-mandal" placeholder="Mandal"></div>
              </div>
              <div class="form-row">
                <div class="form-group"><label>District</label><input type="text" id="reg-district" placeholder="District"></div>
                <div class="form-group"><label>State</label><input type="text" id="reg-state" placeholder="State" value="Telangana"></div>
              </div>
              <div class="form-row">
                <div class="form-group"><label>Pincode</label><input type="text" id="reg-pincode" placeholder="Pincode"></div>
                <div class="form-group"><label>Country</label><input type="text" id="reg-country" placeholder="Country" value="India"></div>
              </div>
              <div class="form-row">
                <div class="form-group"><label>Phone Number *</label><input type="text" id="reg-phone" placeholder="10-digit number" required></div>
                <div class="form-group"><label>Email Address</label><input type="email" id="reg-email" placeholder="patient@gmail.com"></div>
              </div>
            </div>
            <div class="form-section">
              <div class="form-section-title-bar"><h3>Emergency Contact</h3></div>
              <div class="form-row">
                <div class="form-group"><label>Emergency Contact Name</label><input type="text" id="reg-em-name" placeholder="Contact Name"></div>
                <div class="form-group"><label>Relationship</label><input type="text" id="reg-em-relation" placeholder="e.g. Spouse, Son"></div>
              </div>
              <div class="form-row">
                <div class="form-group"><label>Emergency Phone</label><input type="text" id="reg-em-phone" placeholder="Emergency Phone"></div>
                <div class="form-group"><label>Emergency Email</label><input type="email" id="reg-em-email" placeholder="Contact Email"></div>
              </div>
            </div>
            <div style="display:flex;justify-content:flex-end;"><button type="button" class="btn btn-primary" onclick="nextWizardStep()">Next: Medical History ➔</button></div>
          </div>
 
          <!-- STEP 2 -->
          <div id="wizard-step-2" class="wizard-step-content" style="display:none;">
            <div class="form-section">
              <div class="form-section-title-bar"><h3>Referral & Diagnostics</h3></div>
              <div class="form-row">
                <div class="form-group"><label>Referred By</label><select id="reg-ref-by"><option value="Doctor">Doctor</option><option value="Others">Others</option></select></div>
                <div class="form-group"><label>Consultation Type</label><select id="reg-consult-type"><option value="Routine">Routine</option><option value="Emergency">Emergency</option></select></div>
              </div>
              <div class="form-group full-width"><label>Reason For Referral</label><textarea id="reg-ref-reason" placeholder="Reason details"></textarea></div>
              <div class="form-group"><label>Tests Required</label><div class="checkbox-grid"><label class="checkbox-label"><input type="checkbox" id="test-ecg"> ECG</label><label class="checkbox-label"><input type="checkbox" id="test-2de"> 2D ECHO</label><label class="checkbox-label"><input type="checkbox" id="test-tmt"> TMT</label></div></div>
            </div>
            <div class="form-section">
              <div class="form-section-title-bar"><h3>Symptoms & Cardiovascular Profile</h3></div>
              <div class="form-group full-width"><label>Cardiovascular Symptoms</label><textarea id="reg-symptoms-cvs" placeholder="Chest pain, palpitations, dyspnea..."></textarea></div>
              <div class="form-group full-width"><label>Swelling of Feet / Body</label><input type="text" id="reg-swelling" placeholder="Is swelling present?"></div>
              <div class="form-group full-width"><label>Other System Symptoms</label><textarea id="reg-symptoms-other" placeholder="Neurological, gastrointestinal..."></textarea></div>
              <div class="form-group full-width"><label>History of Present Illness</label><textarea id="reg-present-illness" placeholder="Timeline of current symptoms"></textarea></div>
              <div class="form-group full-width"><label>History of Past Illness</label><textarea id="reg-past-illness" placeholder="Previous diagnoses or treatments"></textarea></div>
            </div>
            <div class="form-section">
              <div class="form-section-title-bar"><h3>Risk Factors & Clinical History</h3></div>
              <div class="form-group"><label>Risk Factors</label><div class="checkbox-grid"><label class="checkbox-label"><input type="checkbox" id="risk-dm"> Diabetes (DM)</label><label class="checkbox-label"><input type="checkbox" id="risk-htn"> Hypertension</label><label class="checkbox-label"><input type="checkbox" id="risk-smoke"> Smoking</label><label class="checkbox-label"><input type="checkbox" id="risk-ihd"> Family IHD</label><label class="checkbox-label"><input type="checkbox" id="risk-others"> Others</label></div></div>
              <div class="form-row">
                <div class="form-group"><label>Past Investigations</label><textarea id="reg-past-investigations" placeholder="ECG, blood reports..."></textarea></div>
                <div class="form-group"><label>Hospitalization History</label><textarea id="reg-hospitalization" placeholder="Past hospital stays..."></textarea></div>
              </div>
              <div class="form-row">
                <div class="form-group"><label>Surgery History</label><textarea id="reg-surgery-history" placeholder="Bypass, valve surgeries..."></textarea></div>
                <div class="form-group"><label>Drug / Treatment History</label><textarea id="reg-drug-history" placeholder="Current medications..."></textarea></div>
              </div>
              <div class="form-col-3">
                <div class="form-group"><label>Appetite</label><input type="text" id="reg-appetite" placeholder="Normal/Reduced"></div>
                <div class="form-group"><label>Stools</label><input type="text" id="reg-stools" placeholder="Regular/Constipated"></div>
                <div class="form-group"><label>Urination</label><input type="text" id="reg-urination" placeholder="Normal/Increased"></div>
              </div>
              <div class="form-group full-width"><label>Family History</label><textarea id="reg-family-history" placeholder="Cardiac illness in family..."></textarea></div>
              <div class="form-group full-width"><label>Menstrual & Obstetric History</label><textarea id="reg-menstrual" placeholder="For female patients..."></textarea></div>
              <div class="form-row">
                <div class="form-group"><label>Immunization History</label><textarea id="reg-immunization" placeholder="Covid-19, Influenza..."></textarea></div>
                <div class="form-group"><label>Socioeconomic History</label><textarea id="reg-socioeconomic" placeholder="Lifestyle, diet, stress..."></textarea></div>
              </div>
            </div>
            <div style="display:flex;justify-content:space-between;">
              <button type="button" class="btn btn-secondary" onclick="prevWizardStep()">← Back</button>
              <button type="button" class="btn btn-primary" onclick="nextWizardStep()">Next: Examination ➔</button>
            </div>
          </div>
 
          <!-- STEP 3 -->
          <div id="wizard-step-3" class="wizard-step-content" style="display:none;">
            <div class="form-section">
              <div class="form-section-title-bar"><h3>Biometrics & Vital Parameters</h3></div>
              <div class="form-row">
                <div class="form-group"><label>Height (cm)</label><input type="number" id="reg-height" placeholder="Height in cm" oninput="calculateBMI()"></div>
                <div class="form-group"><label>Weight (kg)</label><input type="number" id="reg-weight" placeholder="Weight in kg" oninput="calculateBMI()"></div>
                <div class="form-group"><label>BMI (Auto-computed)</label><div style="display:flex;gap:8px;"><input type="text" id="reg-bmi" placeholder="BMI" readonly><span class="badge badge-green" style="display:none;" id="reg-bmi-badge">Normal</span></div></div>
              </div>
              <div class="form-row">
                <div class="form-group"><label>Pulse Rate (bpm)</label><input type="text" id="reg-pulse" placeholder="e.g. 72 bpm"></div>
                <div class="form-group"><label>Rhythm</label><input type="text" id="reg-rhythm" placeholder="Regular/Irregular"></div>
                <div class="form-group"><label>Peripheral Pulses</label><input type="text" id="reg-peripheral" placeholder="Normal/Diminished"></div>
              </div>
              <div class="form-col-3">
                <div class="form-group"><label>Supine BP (mmHg)</label><input type="text" id="reg-supine-bp" placeholder="e.g. 120/80"></div>
                <div class="form-group"><label>Sitting BP (mmHg)</label><input type="text" id="reg-sitting-bp" placeholder="e.g. 120/80"></div>
                <div class="form-group"><label>Standing BP (mmHg)</label><input type="text" id="reg-standing-bp" placeholder="e.g. 118/78"></div>
              </div>
            </div>
            <div class="form-section">
              <div class="form-section-title-bar"><h3>CVS Examination</h3></div>
              <div class="form-group full-width"><label>Cardiovascular System (CVS) Exam</label><textarea id="reg-cvs-exam" placeholder="Heart sounds, murmurs..."></textarea></div>
              <div class="form-group full-width"><label>Other Systems Examination</label><textarea id="reg-other-exam" placeholder="Respiratory, CNS, Abdomen..."></textarea></div>
            </div>
            <div class="form-section">
              <div class="form-section-title-bar"><h3>Diagnostic Reports Upload</h3></div>
              <div class="form-row">
                <div class="form-group"><label>Blood Report</label><div class="file-upload-area"><span class="file-upload-icon">🩸</span><div class="file-upload-text">Drag & drop or <span>Browse</span></div><input type="file" id="report-blood" onchange="handleWizardFileUpload(event,'Blood Report')"></div><div class="file-preview-list" id="blood-files-list"></div></div>
                <div class="form-group"><label>ECG Report</label><div class="file-upload-area"><span class="file-upload-icon">📈</span><div class="file-upload-text">Drag & drop or <span>Browse</span></div><input type="file" id="report-ecg" onchange="handleWizardFileUpload(event,'ECG Report')"></div><div class="file-preview-list" id="ecg-files-list"></div></div>
              </div>
              <div class="form-row">
                <div class="form-group"><label>2D ECHO Report</label><div class="file-upload-area"><span class="file-upload-icon">🔊</span><div class="file-upload-text">Drag & drop or <span>Browse</span></div><input type="file" id="report-echo" onchange="handleWizardFileUpload(event,'2D ECHO Report')"></div><div class="file-preview-list" id="echo-files-list"></div></div>
                <div class="form-group"><label>TMT Report</label><div class="file-upload-area"><span class="file-upload-icon">🏃</span><div class="file-upload-text">Drag & drop or <span>Browse</span></div><input type="file" id="report-tmt" onchange="handleWizardFileUpload(event,'TMT Report')"></div><div class="file-preview-list" id="tmt-files-list"></div></div>
              </div>
              <div class="form-row">
                <div class="form-group"><label>CAG (Coronary Angiogram)</label><div class="file-upload-area"><span class="file-upload-icon">🩻</span><div class="file-upload-text">Drag & drop or <span>Browse</span></div><input type="file" id="report-cag" onchange="handleWizardFileUpload(event,'CAG')"></div><div class="file-preview-list" id="cag-files-list"></div></div>
                <div class="form-group"><label>Other Investigations</label><div class="file-upload-area"><span class="file-upload-icon">📁</span><div class="file-upload-text">Drag & drop or <span>Browse</span></div><input type="file" id="report-others" onchange="handleWizardFileUpload(event,'Other Investigations')"></div><div class="file-preview-list" id="others-files-list"></div></div>
              </div>
            </div>
            <div class="form-section">
              <div class="form-section-title-bar"><h3>Management & Final Conclusion</h3></div>
              <div class="form-group full-width"><label>Plan of Management</label><textarea id="reg-management" placeholder="Procedures, dietary plans, follow-up dates..."></textarea></div>
              <div class="form-group full-width"><label>Treatment Advised</label><textarea id="reg-treatment" placeholder="Prescription names and dosages..."></textarea></div>
              <div class="form-group full-width"><label>Conclusion & Recommendations</label><textarea id="reg-conclusion" placeholder="Final cardiac status and general advice..."></textarea></div>
            </div>
            <!-- Excel export note -->
            <div class="excel-banner">
              <div class="excel-banner-text">📊 After saving, patient data will be <strong>automatically saved to localStorage</strong> and you can export all patients to <strong>Excel (.xlsx)</strong> using the Export button in the header.</div>
            </div>
            <div style="display:flex;justify-content:space-between;margin-top:15px;">
              <button type="button" class="btn btn-secondary" onclick="prevWizardStep()">← Back</button>
              <button type="submit" class="btn btn-success" id="save-patient-btn">✅ Save Patient & Send Email</button>
            </div>
          </div>
        </form>
      </div>
    </div>
 
    <!-- OLD PATIENT -->
    <div id="view-old-patient" class="page-view">
      <div class="card">
        <div class="card-header">
          <div class="card-title"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>Srinivasa Patient Register</div>
          <div style="display:flex;gap:10px;align-items:center;">
            <span class="badge badge-green" id="directory-total-badge">0 PATIENTS</span>
            <button class="btn btn-excel" onclick="exportAllPatientsToExcel()" style="padding:6px 12px;font-size:11px;">📊 Export Excel</button>
          </div>
        </div>
        <div class="table-controls">
          <div class="search-bar-wrapper">
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/></svg>
            <input type="text" id="patient-search" placeholder="Search by name, ID, city, or phone..." oninput="handleSearch()">
          </div>
          <div class="filters-wrapper">
            <button class="filter-chip active-chip" id="filter-all" onclick="filterPatients('all')">ALL</button>
            <button class="filter-chip" id="filter-dm" onclick="filterPatients('DM')">DIABETES</button>
            <button class="filter-chip" id="filter-htn" onclick="filterPatients('HTN')">HYPERTENSION</button>
            <button class="filter-chip" id="filter-ihd" onclick="filterPatients('IHD')">FAMILY IHD</button>
          </div>
        </div>
        <div class="table-wrapper">
          <table>
            <thead>
              <tr>
                <th>S.No</th>
                <th>Card No</th>
                <th>Patient Name</th>
                <th>Age/Gender</th>
                <th>Location</th>
                <th>Risk Factors</th>
                <th>Referred By</th>
                <th>Last Updated</th>
                <th style="text-align:center;">Actions</th>
              </tr>
            </thead>
            <tbody id="patient-table-body"></tbody>
          </table>
        </div>
      </div>
    </div>
 
    <!-- APPOINTMENT BOOKING -->
    <div id="view-appointment" class="page-view">
      <div class="grid-2col">
        <div class="card">
          <div class="card-header"><div class="card-title"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>Book Consultation Session</div></div>
          <form id="appointment-booking-form">
            <div class="form-group"><label>Select Registered Patient *</label><select id="book-patient-select" required><option value="">-- Select Patient --</option></select></div>
            <div class="form-row">
              <div class="form-group"><label>Appointment Date *</label><input type="date" id="book-date" onchange="generateTimeSlots()" required></div>
              <div class="form-group"><label>Consultation Type</label><select id="book-consultation"><option value="Routine">Routine (Regular Checkup)</option><option value="Emergency">Emergency Visit</option></select></div>
            </div>
            <div class="form-group"><label>Select Time Slot *</label><div class="scheduler-grid" id="booking-slots-container"></div><input type="hidden" id="selected-time-slot" required></div>
            <div class="form-group"><label>Special Clinical Notes</label><textarea id="book-notes" placeholder="e.g. Chest congestion, palpitation..."></textarea></div>
            <button type="submit" class="btn btn-primary" style="margin-top:10px;width:100%;">Confirm Consultation Booking</button>
          </form>
        </div>
        <div class="card" style="background-color:rgba(11,19,41,0.97);color:white;border:none;">
          <div class="card-header" style="border-bottom-color:rgba(255,255,255,0.08);"><h3 class="card-title" style="color:white;">🕒 Consultation Hours & Info</h3></div>
          <div style="font-size:13px;line-height:1.6;display:flex;flex-direction:column;gap:15px;">
            <p>Welcome to <strong>Srinivasa Heart Centre</strong>. Below are available consultation slots.</p>
            <div class="form-section" style="background:rgba(255,255,255,0.03);border-color:rgba(255,255,255,0.08);padding:15px;">
              <h4 style="color:var(--primary);margin-bottom:8px;">Weekly Schedule</h4>
              <div style="display:flex;justify-content:space-between;margin-bottom:5px;"><span>Mon - Sat:</span><strong>09:00 AM - 05:00 PM</strong></div>
              <div style="display:flex;justify-content:space-between;"><span>Sunday:</span><strong>Emergencies Only</strong></div>
            </div>
          </div>
        </div>
      </div>
    </div>
 
    <!-- APPOINTMENT VIEW -->
    <div id="view-appointment-view" class="page-view">
      <div class="card">
        <div class="card-header">
          <div class="card-title"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 8v4l3 3"/><circle cx="12" cy="12" r="9"/></svg>Clinical Scheduler Board</div>
          <div class="filters-wrapper">
            <button class="filter-chip active-chip" id="appt-filter-today" onclick="filterAppointments('today')">TODAY</button>
            <button class="filter-chip" id="appt-filter-upcoming" onclick="filterAppointments('upcoming')">UPCOMING</button>
            <button class="filter-chip" id="appt-filter-completed" onclick="filterAppointments('completed')">COMPLETED</button>
          </div>
        </div>
        <div class="planner-board" id="appointment-planner-list"></div>
      </div>
    </div>
 
    <!-- STUDY -->
    <div id="view-study" class="page-view">
      <div class="grid-2col">
        <div class="calc-card card">
          <div class="card-header"><h3 class="card-title">❤️ Cardiovascular Risk Assessment</h3></div>
          <p style="font-size:12px;color:var(--text-secondary);line-height:1.4;">Estimate 10-year risk based on clinical metrics.</p>
          <div style="display:flex;flex-direction:column;gap:15px;margin-top:10px;">
            <div class="form-group"><label>Age: <span id="calc-age-val" style="color:var(--primary);">45</span> years</label><input type="range" id="calc-age" min="20" max="80" value="45" class="calc-range-slider" oninput="runRiskCalculator()"></div>
            <div class="form-group"><label>Systolic BP: <span id="calc-sbp-val" style="color:var(--primary);">130</span> mmHg</label><input type="range" id="calc-sbp" min="90" max="200" value="130" class="calc-range-slider" oninput="runRiskCalculator()"></div>
            <div class="form-group"><label>Total Cholesterol: <span id="calc-chol-val" style="color:var(--primary);">200</span> mg/dL</label><input type="range" id="calc-chol" min="100" max="400" value="200" class="calc-range-slider" oninput="runRiskCalculator()"></div>
            <div class="form-group"><label>HDL Cholesterol: <span id="calc-hdl-val" style="color:var(--primary);">45</span> mg/dL</label><input type="range" id="calc-hdl" min="20" max="100" value="45" class="calc-range-slider" oninput="runRiskCalculator()"></div>
            <div class="form-row">
              <div class="form-group"><label>Diabetes</label><select id="calc-dm" onchange="runRiskCalculator()"><option value="no">No</option><option value="yes">Yes</option></select></div>
              <div class="form-group"><label>Smoker</label><select id="calc-smoke" onchange="runRiskCalculator()"><option value="no">No</option><option value="yes">Yes</option></select></div>
            </div>
            <div class="calc-result-box" id="calc-risk-box">
              <div>Estimated 10-Year Cardiovascular Risk:</div>
              <div class="calc-result-value" id="calc-risk-result">4.8%</div>
              <div style="font-weight:700;font-size:11px;text-transform:uppercase;margin-top:5px;" id="calc-risk-status">LOW RISK</div>
            </div>
          </div>
        </div>
        <div style="display:flex;flex-direction:column;gap:25px;">
          <div class="card">
            <div class="card-header"><h3 class="card-title">🏃 Target Heart Rate Calculator</h3></div>
            <div class="form-group"><label>Enter Patient Age</label><input type="number" id="thr-age" value="50" oninput="calculateTHR()"></div>
            <div class="calc-result-box" style="background-color:var(--bg-page);border-color:var(--border-color);">
              <div style="display:flex;justify-content:space-around;">
                <div><label>Max Heart Rate</label><div class="calc-result-value" id="thr-max" style="color:var(--text-primary);">170 bpm</div></div>
                <div style="border-left:1px solid var(--border-color);padding-left:20px;"><label>Aerobic Zone (70-85%)</label><div class="calc-result-value" id="thr-zone" style="color:var(--primary-dark);">119 - 145 bpm</div></div>
              </div>
            </div>
          </div>
          <div class="card">
            <div class="card-header"><h3 class="card-title">📖 JNC-8 Hypertension Reference</h3></div>
            <div style="font-size:12px;display:flex;flex-direction:column;gap:8px;">
              <div style="display:flex;justify-content:space-between;border-bottom:1px solid var(--border-color);padding-bottom:5px;"><strong>Classification</strong><strong>Systolic / Diastolic</strong></div>
              <div style="display:flex;justify-content:space-between;color:var(--success);"><span>Normal BP:</span><span>&lt; 120 and &lt; 80 mmHg</span></div>
              <div style="display:flex;justify-content:space-between;color:var(--warning);"><span>Prehypertension:</span><span>120-139 or 80-89 mmHg</span></div>
              <div style="display:flex;justify-content:space-between;color:var(--accent);"><span>Stage 1 HTN:</span><span>140-159 or 90-99 mmHg</span></div>
              <div style="display:flex;justify-content:space-between;color:#991b1b;font-weight:700;"><span>Stage 2 HTN:</span><span>≥ 160 or ≥ 100 mmHg</span></div>
            </div>
          </div>
        </div>
      </div>
    </div>
 
    <!-- EMAIL -->
    <div id="view-email" class="page-view">
      <div class="email-container">
        <div class="card">
          <div class="card-header"><div class="card-title"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>Compose Email</div></div>
          <form id="email-compose-form">
            <div class="form-row">
              <div class="form-group"><label>Recipient Patient</label><select id="email-patient-select" onchange="autoFillEmailRecipient()" required><option value="">-- Choose Patient --</option></select></div>
              <div class="form-group"><label>Message Template</label><select id="email-template-select" onchange="applyEmailTemplate()"><option value="confirmation">Appointment Confirmation</option><option value="reports">Lab Reports Ready</option><option value="guidelines">Cardiorespiratory Guidelines</option><option value="custom">Custom Email</option></select></div>
            </div>
            <div class="form-group"><label>To (Email Address) *</label><input type="email" id="email-to" required></div>
            <div class="form-group"><label>Subject *</label><input type="text" id="email-subject" required></div>
            <div class="form-group"><label>Email Body</label><textarea id="email-body" style="min-height:200px;" oninput="syncEmailPreview()"></textarea></div>
            <button type="submit" class="btn btn-primary" style="width:100%;" id="send-email-btn">📧 Dispatch Patient Email</button>
          </form>
        </div>
        <div style="display:flex;flex-direction:column;gap:15px;">
          <div style="font-size:12px;font-weight:700;color:var(--text-secondary);text-transform:uppercase;">Live Email Preview</div>
          <div class="email-preview-frame">
            <div class="email-preview-header">
              <div style="width:34px;height:34px;border-radius:50%;background-color:#ecfeff;display:flex;align-items:center;justify-content:center;color:#06b6d4;font-weight:700;">❤️</div>
              <div><h4 style="margin:0;font-family:'Outfit',sans-serif;">Srinivasa Heart Centre</h4><small style="color:#64748b;">Dr. Srinivas Ramaka, MD, DM.</small></div>
            </div>
            <div style="font-size:12px;color:#475569;border-bottom:1px solid #f1f5f9;padding-bottom:10px;"><strong>Subject:</strong> <span id="preview-subj">Appointment Confirmation</span></div>
            <div class="email-preview-body" id="preview-body-content"></div>
            <div class="email-preview-footer"><p>Srinivasa Heart Centre | Warangal, Telangana, India.</p></div>
          </div>
        </div>
      </div>
    </div>
 
    <!-- ACCOUNTS -->
    <div id="view-accounts" class="page-view">
      <div class="stats-grid">
        <div class="stat-card"><div class="stat-info"><h3>Total Billings</h3><div class="stat-value" id="billing-stat-total">₹0</div></div><div class="stat-icon">💰</div></div>
        <div class="stat-card"><div class="stat-info"><h3>Paid Invoice Value</h3><div class="stat-value" id="billing-stat-paid" style="color:var(--success);">₹0</div></div><div class="stat-icon">✅</div></div>
        <div class="stat-card"><div class="stat-info"><h3>Outstanding Balance</h3><div class="stat-value" id="billing-stat-pending" style="color:var(--accent);">₹0</div></div><div class="stat-icon">⚠️</div></div>
      </div>
      <div class="grid-2col">
        <div class="card">
          <div class="card-header"><div class="card-title"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>Generate Invoice</div></div>
          <form id="billing-invoice-form">
            <div class="form-group"><label>Select Patient</label><select id="bill-patient-select" required><option value="">-- Choose Patient --</option></select></div>
            <div class="form-group">
              <label>Diagnostic Services & Fees</label>
              <div class="invoice-builder-items">
                <label class="checkbox-label" style="display:flex;justify-content:space-between;"><span><input type="checkbox" class="bill-item-chk" value="1000" data-name="Consultation Fee" checked> Consultation Fee</span><strong>₹1,000</strong></label>
                <label class="checkbox-label" style="display:flex;justify-content:space-between;"><span><input type="checkbox" class="bill-item-chk" value="500" data-name="ECG Report Analysis"> ECG Report Analysis</span><strong>₹500</strong></label>
                <label class="checkbox-label" style="display:flex;justify-content:space-between;"><span><input type="checkbox" class="bill-item-chk" value="2500" data-name="2D Echocardiogram"> 2D Echocardiogram</span><strong>₹2,500</strong></label>
                <label class="checkbox-label" style="display:flex;justify-content:space-between;"><span><input type="checkbox" class="bill-item-chk" value="2000" data-name="Treadmill Test (TMT)"> Treadmill Test (TMT)</span><strong>₹2,000</strong></label>
              </div>
            </div>
            <div class="form-row">
              <div class="form-group"><label>Payment Method</label>
                <select id="bill-payment-method" onchange="togglePaymentQR()">
                  <option value="Cash">💵 Cash</option>
                  <option value="Card">💳 Debit/Credit Card</option>
                  <option value="UPI - PhonePe">📱 UPI - PhonePe</option>
                  <option value="UPI - Google Pay">📱 UPI - Google Pay</option>
                  <option value="UPI - Paytm">📱 UPI - Paytm</option>
                  <option value="Net Banking">🏦 Net Banking</option>
                </select>
              </div>
              <div class="form-group"><label>Payment Status</label><select id="bill-invoice-status"><option value="Paid">Paid</option><option value="Pending">Pending / Unpaid</option></select></div>
            </div>
            <div id="upi-qr-box" style="display:none;text-align:center;border:1px dashed var(--border-color);background-color:var(--bg-page);border-radius:var(--radius-sm);padding:16px;margin-bottom:15px;">
              <div style="font-size:11px;font-weight:700;color:var(--text-secondary);text-transform:uppercase;margin-bottom:10px;">Scan to Pay</div>
              <div id="upi-qr-canvas-wrapper" style="display:flex;justify-content:center;"></div>
              <div style="font-size:11px;color:var(--text-muted);margin-top:8px;">UPI ID: <strong id="upi-id-display">srinivasaheartcentre@upi</strong></div>
            </div>
            <div class="form-row">
              <div class="form-group"><label>Invoice Total</label><div class="calc-result-value" id="bill-running-total" style="font-size:20px;padding:6px 0;">₹1,000</div></div>
            </div>
            <button type="submit" class="btn btn-primary" style="width:100%;">Issue Invoice & Email to Patient</button>
          </form>
        </div>
        <div class="card">
          <div class="card-header"><h3 class="card-title">📋 Invoice Ledger</h3></div>
          <div class="table-wrapper">
            <table>
              <thead><tr><th>Patient</th><th>Items</th><th>Total</th><th>Method</th><th>Status</th><th>Action</th></tr></thead>
              <tbody id="billing-ledger-body"></tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
 
  </div>
</div>
 
<!-- PATIENT VIEW MODAL -->
<div class="modal-overlay" id="patient-details-modal">
  <div class="modal-content">
    <div class="modal-header">
      <h3 class="modal-title" id="modal-title-name">Clinical Case File</h3>
      <button class="btn btn-secondary btn-icon" onclick="closeModal('patient-details-modal')">
        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M6 18L18 6M6 6l12 12"/></svg>
      </button>
    </div>
    <div class="modal-body" id="modal-body-content"></div>
    <div class="modal-footer">
      <button class="btn btn-secondary" onclick="printClinicalFile()">🖨️ Print</button>
      <button class="btn btn-primary" onclick="closeModal('patient-details-modal')">Close</button>
    </div>
  </div>
</div>
 
<!-- EDIT PATIENT MODAL -->
<div class="modal-overlay" id="edit-patient-modal">
  <div class="modal-content edit-modal-content">
    <div class="modal-header">
      <h3 class="modal-title" id="edit-modal-title">✏️ Edit Patient Details</h3>
      <button class="btn btn-secondary btn-icon" onclick="closeModal('edit-patient-modal')">
        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M6 18L18 6M6 6l12 12"/></svg>
      </button>
    </div>
    <div class="modal-body" id="edit-modal-body">
      <form id="edit-patient-form">
        <!-- PERSONAL -->
        <div class="form-section">
          <div class="form-section-title-bar"><h3>Personal Details</h3></div>
          <div class="form-row">
            <div class="form-group"><label>First Name *</label><input type="text" id="edit-first-name" required></div>
            <div class="form-group"><label>Last Name *</label><input type="text" id="edit-last-name" required></div>
          </div>
          <div class="form-row">
            <div class="form-group"><label>Date of Birth</label><input type="date" id="edit-dob" onchange="calculateEditAge()"></div>
            <div class="form-group"><label>Age</label><input type="number" id="edit-age" min="0" max="120"></div>
            <div class="form-group"><label>Gender</label><select id="edit-gender"><option value="Male">Male</option><option value="Female">Female</option><option value="Other">Other</option></select></div>
          </div>
          <div class="form-row">
            <div class="form-group"><label>Phone *</label><input type="text" id="edit-phone" required></div>
            <div class="form-group"><label>Email</label><input type="email" id="edit-email"></div>
            <div class="form-group"><label>Occupation</label><input type="text" id="edit-occupation"></div>
          </div>
          <div class="form-group full-width"><label>Address</label><textarea id="edit-address"></textarea></div>
          <div class="form-row">
            <div class="form-group"><label>Home Town</label><input type="text" id="edit-hometown"></div>
            <div class="form-group"><label>District</label><input type="text" id="edit-district"></div>
            <div class="form-group"><label>State</label><input type="text" id="edit-state"></div>
          </div>
        </div>
 
        <!-- MEDICAL -->
        <div class="form-section">
          <div class="form-section-title-bar"><h3>Medical History & Symptoms</h3></div>
          <div class="form-row">
            <div class="form-group"><label>Referred By</label><select id="edit-ref-by"><option value="Doctor">Doctor</option><option value="Others">Others</option></select></div>
            <div class="form-group"><label>Consultation Type</label><select id="edit-consult-type"><option value="Routine">Routine</option><option value="Emergency">Emergency</option></select></div>
          </div>
          <div class="form-group full-width"><label>Referral Reason</label><textarea id="edit-ref-reason"></textarea></div>
          <div class="form-group full-width"><label>Cardiovascular Symptoms</label><textarea id="edit-cvs-symptoms"></textarea></div>
          <div class="form-group full-width"><label>Present Illness History</label><textarea id="edit-present-illness"></textarea></div>
          <div class="form-group"><label>Risk Factors</label>
            <div class="checkbox-grid">
              <label class="checkbox-label"><input type="checkbox" id="edit-risk-dm"> Diabetes (DM)</label>
              <label class="checkbox-label"><input type="checkbox" id="edit-risk-htn"> Hypertension</label>
              <label class="checkbox-label"><input type="checkbox" id="edit-risk-smoke"> Smoking</label>
              <label class="checkbox-label"><input type="checkbox" id="edit-risk-ihd"> Family IHD</label>
              <label class="checkbox-label"><input type="checkbox" id="edit-risk-others"> Others</label>
            </div>
          </div>
          <div class="form-group full-width"><label>Drug / Treatment History</label><textarea id="edit-drug-history"></textarea></div>
          <div class="form-group full-width"><label>Family History</label><textarea id="edit-family-history"></textarea></div>
        </div>
 
        <!-- VITALS -->
        <div class="form-section">
          <div class="form-section-title-bar"><h3>Vitals & Examination</h3></div>
          <div class="form-row">
            <div class="form-group"><label>Height (cm)</label><input type="number" id="edit-height" oninput="calculateEditBMI()"></div>
            <div class="form-group"><label>Weight (kg)</label><input type="number" id="edit-weight" oninput="calculateEditBMI()"></div>
            <div class="form-group"><label>BMI (Auto)</label><input type="text" id="edit-bmi" readonly></div>
          </div>
          <div class="form-row">
            <div class="form-group"><label>Pulse Rate</label><input type="text" id="edit-pulse"></div>
            <div class="form-group"><label>Supine BP</label><input type="text" id="edit-supine-bp"></div>
            <div class="form-group"><label>Sitting BP</label><input type="text" id="edit-sitting-bp"></div>
          </div>
          <div class="form-group full-width"><label>CVS Examination</label><textarea id="edit-cvs-exam"></textarea></div>
        </div>
 
        <!-- CONCLUSION -->
        <div class="form-section">
          <div class="form-section-title-bar"><h3>Conclusion & Treatment</h3></div>
          <div class="form-group full-width"><label>Plan of Management</label><textarea id="edit-management"></textarea></div>
          <div class="form-group full-width"><label>Treatment Advised</label><textarea id="edit-treatment"></textarea></div>
          <div class="form-group full-width"><label>Conclusion & Recommendations</label><textarea id="edit-conclusion"></textarea></div>
        </div>
 
        <div style="display:flex;justify-content:flex-end;gap:10px;">
          <button type="button" class="btn btn-secondary" onclick="closeModal('edit-patient-modal')">Cancel</button>
          <button type="submit" class="btn btn-primary">💾 Save Changes & Update Excel</button>
        </div>
      </form>
    </div>
  </div>
</div>
 
<div class="toast-container" id="toast-alerts-box"></div>
 
<script>
// ===== EMAILJS CONFIGURATION =====
const EMAILJS_CONFIG = {
  PUBLIC_KEY: "gGWN5pFl2k8IBzDmC",
  SERVICE_ID: "service_8h5j02l",
  TEMPLATE_ID: "template_260ejiw"
};

// Initialize EmailJS
(function() {
  emailjs.init({
    publicKey: EMAILJS_CONFIG.PUBLIC_KEY
  });
  console.log("✅ EmailJS initialized successfully!");
  console.log("📋 Template ID:", EMAILJS_CONFIG.TEMPLATE_ID);
})();

// ===== FUNCTION TO SEND CONFIRMATION EMAIL =====
// FIX: now trims the email, validates it's non-empty BEFORE calling EmailJS,
// and sends several common recipient-variable spellings (to_email / email /
// recipient / reply_to) so it matches whatever variable name your EmailJS
// "To Email" field on the Service/Template is bound to.
function sendPatientEmail(patientData) {
  const patientName = (patientData.firstName + " " + patientData.lastName).trim();
  const patientEmail = (patientData.email || "").trim();
  const patientPhone = patientData.phone;

  // If no email is provided, skip sending (this is expected/normal, not an error)
  if (!patientEmail) {
    console.log("ℹ️ No email provided for patient. Skipping email.");
    return Promise.resolve("No email provided");
  }

  console.log("📧 Sending email to:", patientEmail);
  console.log("📋 Patient Name:", patientName);
  console.log("📋 Patient Phone:", patientPhone);

  // IMPORTANT: EmailJS reads the recipient address from whichever variable
  // name you typed into the "To Email" box on your EmailJS Service/Template
  // settings page (NOT from the message body). We send the same address
  // under several common names so it matches regardless of which one your
  // template actually uses.
  const templateParams = {
    to_email: patientEmail,
    email: patientEmail,
    recipient: patientEmail,
    reply_to: patientEmail,
    to_name: patientName,
    patient_phone: patientPhone,
    hospital_name: "Srinivasa Heart Centre",
    message: "Your registration has been completed successfully. Welcome to Srinivasa Heart Centre, Dr. Srinivas Ramaka, MD, DM."
  };

  console.log("📋 Sending with params:", templateParams);

  return emailjs.send(
    EMAILJS_CONFIG.SERVICE_ID,
    EMAILJS_CONFIG.TEMPLATE_ID,
    templateParams
  )
  .then(function(response) {
    console.log("✅ Confirmation email sent successfully:", response);
    showToast(`✅ Confirmation email sent to ${patientEmail}!`, "success");
    return response;
  })
  .catch(function(error) {
    console.error("❌ Email sending failed:", error);
    const errText = (error && (error.text || error.message)) || "Unknown error";
    if(/recipient/i.test(errText) && /empty/i.test(errText)){
      showToast(`⚠️ Patient saved, but EmailJS rejected the recipient. Open your EmailJS template settings and set "To Email" to {{to_email}}.`, "error");
    } else {
      showToast(`⚠️ Patient saved but email not sent: ${errText}`, "error");
    }
    return error;
  });
}

// ===== DATA =====
const defaultPatients = [
  {id:"PT-708",cardType:"Aadhar",cardNumber:"4820-2910-3847",firstName:"Venkatesh",lastName:"Rao",dob:"1968-04-12",age:58,gender:"Male",photo:null,occupation:"Retired Bank Manager",address:"Lane 4, Subedari",hometown:"Hanamkonda",mandal:"Hanamkonda",district:"Warangal",state:"Telangana",pincode:"506001",country:"India",phone:"9848022338",email:"venkat.rao68@gmail.com",emName:"Sudha Rao",emRelation:"Spouse",emPhone:"9440382910",emEmail:"sudha.rao@gmail.com",refBy:"Doctor",refReason:"Frequent breathlessness, stage II HTN.",consultType:"Routine",testsRequired:{ecg:true,echo:true,tmt:false},cvsSymptoms:"Chest tightness radiating to left arm.",swelling:"Bilateral pedal edema, mild (+1)",otherSymptoms:"No dizziness.",presentIllness:"Dyspnea on walking uphill for 2 months.",pastIllness:"Hypertension 5 years ago.",riskFactors:{dm:true,htn:true,smoke:false,ihd:true,others:false},pastInvestigations:"ECG shows inverted T waves. HbA1c: 7.2%",hospitalization:"None.",surgeryHistory:"None.",drugHistory:"Amlodipine 5mg OD.",appetite:"Normal",stools:"Regular",urination:"Normal",familyHistory:"Father died of MI at 62.",menstrual:"",immunization:"Fully vaccinated.",socioeconomic:"Low-stress retirement.",height:170,weight:78,bmi:"27.0",pulse:"82 bpm",rhythm:"Regular",peripheral:"All palpable",supineBP:"150/92",sittingBP:"148/90",standingBP:"142/88",cvsExam:"S1 S2 heard. No murmurs.",otherExam:"Abdomen soft. CNS intact.",files:["ECG_Report.pdf"],management:"Advised CAG. Start dual antiplatelets.",treatment:"Tab Ecosprin 150mg OD\nTab Clopidogrel 75mg OD\nTab Atorvastatin 40mg HS",conclusion:"Suspected Stable Angina. CAD evaluation suggested.",lastEdited:null,editHistory:[]},
  {id:"PT-205",cardType:"PAN",cardNumber:"BPYPR1048L",firstName:"Rajitha",lastName:"Reddy",dob:"1979-08-25",age:46,gender:"Female",photo:null,occupation:"School Teacher",address:"Flat 202, Gokul Apartments",hometown:"Warangal",mandal:"Warangal",district:"Warangal",state:"Telangana",pincode:"506002",country:"India",phone:"7702938471",email:"rajitha.teacher@gmail.com",emName:"Karan Reddy",emRelation:"Husband",emPhone:"9000182938",emEmail:"karan.reddy@gmail.com",refBy:"Doctor",refReason:"Abnormal ECG findings.",consultType:"Routine",testsRequired:{ecg:true,echo:true,tmt:true},cvsSymptoms:"Occasional palpitations when resting.",swelling:"No swelling.",otherSymptoms:"Mild anxiety.",presentIllness:"Palpitations for 6 months, self-limiting.",pastIllness:"Gestational diabetes 2012, resolved.",riskFactors:{dm:false,htn:false,smoke:false,ihd:false,others:true},pastInvestigations:"Normal echo 2 years ago.",hospitalization:"None.",surgeryHistory:"Cesarean section 2012.",drugHistory:"None.",appetite:"Good",stools:"Regular",urination:"Normal",familyHistory:"No cardiac history.",menstrual:"Regular periods.",immunization:"Vaccinated.",socioeconomic:"Moderate job stress.",height:162,weight:58,bmi:"22.1",pulse:"88 bpm",rhythm:"Irregular (PVCs)",peripheral:"Palpable",supineBP:"118/76",sittingBP:"116/74",standingBP:"114/72",cvsExam:"Frequent PVCs. No murmurs.",otherExam:"Clear respiratory system.",files:["ECG_Holter.pdf"],management:"Holter monitoring advised.",treatment:"Tab Metoprolol 25mg OD\nReduce caffeine",conclusion:"Frequent PVCs on structurally normal heart.",lastEdited:null,editHistory:[]}
];
const defaultAppointments = [
  {id:"AP-101",patientId:"PT-708",date:new Date().toISOString().split('T')[0],time:"09:30 AM",consultType:"Routine",notes:"Cardiology consultation",status:"Checked-In"},
  {id:"AP-102",patientId:"PT-205",date:new Date().toISOString().split('T')[0],time:"11:00 AM",consultType:"Routine",notes:"Palpitation review",status:"Scheduled"}
];
const defaultInvoices = [
  {id:"INV-101",patientId:"PT-708",date:new Date().toISOString().split('T')[0],items:["Consultation Fee","ECG Report Analysis","2D Echocardiogram"],amount:4000,status:"Paid"},
  {id:"INV-102",patientId:"PT-205",date:new Date().toISOString().split('T')[0],items:["Consultation Fee"],amount:1000,status:"Pending"}
];
 
let DB_PATIENTS = JSON.parse(localStorage.getItem("SHC_PATIENTS")) || defaultPatients;
let DB_APPOINTMENTS = JSON.parse(localStorage.getItem("SHC_APPOINTMENTS")) || defaultAppointments;
let DB_INVOICES = JSON.parse(localStorage.getItem("SHC_INVOICES")) || defaultInvoices;
 
function savePatients(){ localStorage.setItem("SHC_PATIENTS", JSON.stringify(DB_PATIENTS)); }
function saveAppointments(){ localStorage.setItem("SHC_APPOINTMENTS", JSON.stringify(DB_APPOINTMENTS)); }
function saveInvoices(){ localStorage.setItem("SHC_INVOICES", JSON.stringify(DB_INVOICES)); }
 
let currentWizardStep = 1;
let uploadedFilesStore = {"Blood Report":[],"ECG Report":[],"2D ECHO Report":[],"TMT Report":[],"CAG":[],"Other Investigations":[]};
let registeredPhotoData = null;
let currentApptFilter = "today";
let currentPatientFilter = "all";
let dashboardChartInstance = null;
let editingPatientId = null;
 
// ===== DATE / EDIT-HISTORY HELPERS =====
function formatDateTime(iso){
  if(!iso) return "";
  const d = new Date(iso);
  if(isNaN(d.getTime())) return "";
  return d.toLocaleString('en-IN',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit',hour12:true});
}
 
const EDITABLE_FIELD_LABELS = {
  firstName:"First Name", lastName:"Last Name", dob:"Date of Birth", age:"Age", gender:"Gender",
  phone:"Phone", email:"Email", occupation:"Occupation", address:"Address", hometown:"Home Town",
  district:"District", state:"State", refBy:"Referred By", consultType:"Consultation Type",
  refReason:"Referral Reason", cvsSymptoms:"CVS Symptoms", presentIllness:"Present Illness",
  drugHistory:"Drug History", familyHistory:"Family History", height:"Height (cm)", weight:"Weight (kg)",
  bmi:"BMI", pulse:"Pulse Rate", supineBP:"Supine BP", sittingBP:"Sitting BP", cvsExam:"CVS Exam",
  management:"Management Plan", treatment:"Treatment", conclusion:"Conclusion"
};
const EDITABLE_RISK_LABELS = {dm:"Diabetes (DM)",htn:"Hypertension",smoke:"Smoking",ihd:"Family IHD",others:"Other Risk Factors"};
 
function buildEditDiff(oldP, newVals){
  const changes=[];
  Object.keys(EDITABLE_FIELD_LABELS).forEach(key=>{
    const oldVal = (oldP[key]===undefined||oldP[key]===null) ? "" : String(oldP[key]);
    const newVal = (newVals[key]===undefined||newVals[key]===null) ? "" : String(newVals[key]);
    if(oldVal!==newVal){
      changes.push({field:EDITABLE_FIELD_LABELS[key], oldValue:oldVal||"—", newValue:newVal||"—"});
    }
  });
  Object.keys(EDITABLE_RISK_LABELS).forEach(key=>{
    const oldVal = oldP.riskFactors ? !!oldP.riskFactors[key] : false;
    const newVal = newVals.riskFactors ? !!newVals.riskFactors[key] : false;
    if(oldVal!==newVal){
      changes.push({field:EDITABLE_RISK_LABELS[key], oldValue:oldVal?"Yes":"No", newValue:newVal?"Yes":"No"});
    }
  });
  return changes;
}
 
function buildEditHistoryHtml(p){
  if(!p.editHistory || !p.editHistory.length){
    return `<p style="color:var(--text-muted);font-size:12px;">No edits made since registration.</p>`;
  }
  return [...p.editHistory].reverse().map(entry=>{
    const rows = entry.changes.map(c=>`
      <div style="display:flex;flex-wrap:wrap;justify-content:space-between;gap:10px;padding:7px 0;border-bottom:1px solid var(--border-color);font-size:12px;">
        <strong style="min-width:140px;">${c.field}</strong>
        <span style="color:var(--accent);text-decoration:line-through;flex:1;min-width:100px;">${c.oldValue}</span>
        <span style="color:var(--text-muted);">➔</span>
        <span style="color:var(--success);flex:1;min-width:100px;font-weight:600;">${c.newValue}</span>
      </div>`).join("");
    return `<div style="margin-bottom:14px;background:var(--bg-page);border:1px solid var(--border-color);border-radius:8px;padding:12px;">
      <div style="font-weight:700;font-size:12px;color:var(--primary);margin-bottom:6px;">🕒 Edited on ${formatDateTime(entry.date)}</div>
      ${rows}
    </div>`;
  }).join("");
}
 
// ===== CLOCK =====
function initClock(){
  function tick(){
    document.getElementById('clock-display').innerHTML = new Date().toLocaleString('en-US',{weekday:'short',year:'numeric',month:'short',day:'numeric',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:true});
  }
  setInterval(tick,1000); tick();
}
 
// ===== TOAST =====
function showToast(message, type="success"){
  const box = document.getElementById("toast-alerts-box");
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  const icons = {success:"✅",error:"❌",info:"ℹ️"};
  toast.innerHTML = `<div>${icons[type]||"🔔"}</div><div>${message}</div>`;
  box.appendChild(toast);
  setTimeout(()=>{ toast.style.opacity="0"; toast.style.transition="opacity 0.3s"; setTimeout(()=>toast.remove(),300); }, 4000);
}
 
// ===== EXCEL EXPORT =====
function exportAllPatientsToExcel(){
  if(DB_PATIENTS.length === 0){
    showToast("No patients to export!", "error");
    return;
  }
 
  const wb = XLSX.utils.book_new();
 
  const patientRows = DB_PATIENTS.map((p,i) => ({
    "S.No": i+1,
    "Patient ID": p.id,
    "First Name": p.firstName,
    "Last Name": p.lastName,
    "Full Name": `${p.firstName} ${p.lastName}`,
    "Date of Birth": p.dob,
    "Age": p.age,
    "Gender": p.gender,
    "Occupation": p.occupation || "",
    "Phone": p.phone,
    "Email": p.email || "",
    "Address": p.address || "",
    "Home Town": p.hometown || "",
    "District": p.district || "",
    "State": p.state || "",
    "Pincode": p.pincode || "",
    "Card Type": p.cardType,
    "Card Number": p.cardNumber,
    "Emergency Contact": p.emName || "",
    "Emergency Phone": p.emPhone || "",
    "Referred By": p.refBy,
    "Consultation Type": p.consultType,
    "Referral Reason": p.refReason || "",
    "CVS Symptoms": p.cvsSymptoms || "",
    "Present Illness": p.presentIllness || "",
    "Past Illness": p.pastIllness || "",
    "Diabetes (DM)": p.riskFactors.dm ? "Yes" : "No",
    "Hypertension (HTN)": p.riskFactors.htn ? "Yes" : "No",
    "Smoking": p.riskFactors.smoke ? "Yes" : "No",
    "Family IHD": p.riskFactors.ihd ? "Yes" : "No",
    "Drug History": p.drugHistory || "",
    "Height (cm)": p.height || "",
    "Weight (kg)": p.weight || "",
    "BMI": p.bmi || "",
    "Pulse Rate": p.pulse || "",
    "Supine BP": p.supineBP || "",
    "Sitting BP": p.sittingBP || "",
    "CVS Exam": p.cvsExam || "",
    "Management Plan": p.management || "",
    "Treatment": p.treatment || "",
    "Conclusion": p.conclusion || "",
    "Last Edited": p.lastEdited ? formatDateTime(p.lastEdited) : "Not edited",
    "Total Edits": p.editHistory ? p.editHistory.length : 0,
  }));
 
  const ws1 = XLSX.utils.json_to_sheet(patientRows);
  ws1['!cols'] = [
    {wch:6},{wch:10},{wch:14},{wch:14},{wch:22},{wch:14},{wch:6},{wch:8},{wch:18},
    {wch:14},{wch:24},{wch:24},{wch:14},{wch:14},{wch:12},{wch:8},
    {wch:10},{wch:18},{wch:18},{wch:14},{wch:12},{wch:14},{wch:30},
    {wch:30},{wch:30},{wch:30},{wch:10},{wch:12},{wch:10},{wch:12},
    {wch:30},{wch:10},{wch:10},{wch:8},{wch:12},{wch:12},{wch:12},
    {wch:30},{wch:30},{wch:35},{wch:35},{wch:20},{wch:10}
  ];
  XLSX.utils.book_append_sheet(wb, ws1, "Patients");
 
  const apptRows = DB_APPOINTMENTS.map(a => {
    const p = DB_PATIENTS.find(x => x.id === a.patientId);
    return {
      "Appointment ID": a.id,
      "Patient ID": a.patientId,
      "Patient Name": p ? `${p.firstName} ${p.lastName}` : "Unknown",
      "Date": a.date,
      "Time": a.time,
      "Type": a.consultType,
      "Notes": a.notes || "",
      "Status": a.status
    };
  });
  const ws2 = XLSX.utils.json_to_sheet(apptRows);
  ws2['!cols'] = [{wch:12},{wch:10},{wch:22},{wch:12},{wch:10},{wch:12},{wch:30},{wch:12}];
  XLSX.utils.book_append_sheet(wb, ws2, "Appointments");
 
  const invRows = DB_INVOICES.map(inv => {
    const p = DB_PATIENTS.find(x => x.id === inv.patientId);
    return {
      "Invoice ID": inv.id,
      "Patient ID": inv.patientId,
      "Patient Name": p ? `${p.firstName} ${p.lastName}` : "Unknown",
      "Date": inv.date,
      "Services": inv.items.join(", "),
      "Amount (₹)": inv.amount,
      "Status": inv.status
    };
  });
  const ws3 = XLSX.utils.json_to_sheet(invRows);
  ws3['!cols'] = [{wch:12},{wch:10},{wch:22},{wch:12},{wch:40},{wch:12},{wch:10}];
  XLSX.utils.book_append_sheet(wb, ws3, "Invoices");
 
  const historyRows = [];
  DB_PATIENTS.forEach(p=>{
    (p.editHistory||[]).forEach(entry=>{
      entry.changes.forEach(c=>{
        historyRows.push({
          "Patient ID": p.id,
          "Patient Name": `${p.firstName} ${p.lastName}`,
          "Edited On": formatDateTime(entry.date),
          "Field Changed": c.field,
          "Old Value": c.oldValue,
          "New Value": c.newValue
        });
      });
    });
  });
  const ws4 = XLSX.utils.json_to_sheet(historyRows.length ? historyRows : [{"Patient ID":"","Patient Name":"","Edited On":"","Field Changed":"","Old Value":"","New Value":""}]);
  ws4['!cols'] = [{wch:10},{wch:22},{wch:20},{wch:20},{wch:30},{wch:30}];
  XLSX.utils.book_append_sheet(wb, ws4, "Edit History");
 
  const filename = `Srinivasa_Heart_Centre_${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(wb, filename);
  showToast(`Excel exported: ${filename}`, "success");
}
 
// ===== NAVIGATION =====
function switchView(viewId){
  document.querySelectorAll(".sidebar-item").forEach(i => i.classList.remove("active"));
  document.querySelectorAll(".page-view").forEach(p => p.classList.remove("active-view"));
  const tm = document.getElementById(`menu-${viewId}`);
  if(tm) tm.classList.add("active");
  const tp = document.getElementById(`view-${viewId}`);
  if(tp) tp.classList.add("active-view");
 
  const titles = {
    home:["Srinivasa Heart Centre","Consultant Cardiologist Diagnostic Dashboard"],
    "new-patient":["New Patient Enrollment","Register personal details, medical history and reports"],
    "old-patient":["Registered Patients Registry","View, edit, and manage patient profiles"],
    appointment:["Schedule Consultation","Book a new clinical appointment"],
    "appointment-view":["Consultation Schedule","Monitor daily schedule board"],
    study:["Clinical Calculators","Interactive cardiac guidelines tools"],
    email:["Patient Email Portal","Send notifications to patients"],
    accounts:["Billing & Accounts","Invoicing and payment management"]
  };
  const t = titles[viewId] || ["Srinivasa Heart Centre","Dashboard"];
  document.getElementById("page-title-display").innerText = t[0];
  document.getElementById("page-subtitle-display").innerText = t[1];
 
  if(viewId==="home") renderDashboard();
  else if(viewId==="new-patient"){ resetWizard(); }
  else if(viewId==="old-patient"){ renderPatientDirectory(); }
  else if(viewId==="appointment"){ populatePatientDropdowns(); generateTimeSlots(); }
  else if(viewId==="appointment-view"){ renderAppointmentsPlanner(); }
  else if(viewId==="study"){ runRiskCalculator(); calculateTHR(); }
  else if(viewId==="email"){ populatePatientDropdowns(); applyEmailTemplate(); }
  else if(viewId==="accounts"){ populatePatientDropdowns(); renderBillingLedger(); }
}
 
// ===== DASHBOARD =====
function renderDashboard(){
  document.getElementById("stat-total-patients").innerText = DB_PATIENTS.length;
  const today = new Date().toISOString().split('T')[0];
  document.getElementById("stat-today-appointments").innerText = DB_APPOINTMENTS.filter(a=>a.date===today).length;
  const rev = DB_INVOICES.reduce((a,b)=>a+b.amount,0);
  document.getElementById("stat-collected-revenue").innerText = `₹${rev.toLocaleString('en-IN')}`;
 
  const ac = document.getElementById("dashboard-recent-activity");
  ac.innerHTML = "";
  [...DB_PATIENTS].reverse().slice(0,4).forEach(p=>{
    const init = `${p.firstName.charAt(0)}${p.lastName.charAt(0)}`.toUpperCase();
    const avatar = p.photo
      ? `<div style="width:34px;height:34px;border-radius:50%;background-image:url('${p.photo}');background-size:cover;"></div>`
      : `<div style="width:34px;height:34px;border-radius:50%;background-color:var(--primary-light);color:var(--primary);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;border:1px solid var(--primary);">${init}</div>`;
    const d = document.createElement("div");
    d.style="display:flex;align-items:center;justify-content:space-between;padding:12px;background:var(--bg-page);border-radius:var(--radius-sm);border:1px solid var(--border-color);";
    d.innerHTML=`<div style="display:flex;align-items:center;gap:10px;">${avatar}<div><h4 style="font-size:13px;font-weight:700;">${p.firstName} ${p.lastName}</h4><small style="color:var(--text-secondary);font-size:11px;">${p.id} • ${p.gender}, ${p.age} yrs</small></div></div><span class="badge badge-blue" style="font-size:9px;">Registered</span>`;
    ac.appendChild(d);
  });
  initDashboardChart();
}
 
function initDashboardChart(){
  const ctx = document.getElementById('dashboardChart').getContext('2d');
  let dm=0,htn=0,smoke=0,ihd=0;
  DB_PATIENTS.forEach(p=>{ if(p.riskFactors.dm)dm++; if(p.riskFactors.htn)htn++; if(p.riskFactors.smoke)smoke++; if(p.riskFactors.ihd)ihd++; });
  if(dashboardChartInstance) dashboardChartInstance.destroy();
  const isDark = document.documentElement.getAttribute("data-theme")==="dark";
  dashboardChartInstance = new Chart(ctx, {
    type:'bar',
    data:{labels:['Diabetes','Hypertension','Smoking','Family IHD'],datasets:[{label:'Risk Factors',data:[dm,htn,smoke,ihd],backgroundColor:['rgba(6,182,212,0.65)','rgba(239,68,68,0.65)','rgba(245,158,11,0.65)','rgba(16,185,129,0.65)'],borderColor:['#06b6d4','#ef4444','#f59e0b','#10b981'],borderWidth:1.5,borderRadius:6}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{color:isDark?"#94a3b8":"#475569",font:{size:10}},grid:{color:isDark?"rgba(255,255,255,0.05)":"#e2e8f0"}},x:{ticks:{color:isDark?"#94a3b8":"#475569",font:{size:10}},grid:{display:false}}}}
  });
}
 
// ===== WIZARD =====
function resetWizard(){
  currentWizardStep=1;
  uploadedFilesStore={"Blood Report":[],"ECG Report":[],"2D ECHO Report":[],"TMT Report":[],"CAG":[],"Other Investigations":[]};
  registeredPhotoData=null;
  document.getElementById("wizard-patient-form").reset();
  document.getElementById("photo-upload-preview").innerText="📸";
  document.getElementById("photo-upload-preview").style.backgroundImage="none";
  document.querySelectorAll(".file-preview-list").forEach(l=>l.innerHTML="");
  document.querySelectorAll(".checkbox-label").forEach(l=>l.classList.remove("checked"));
  goToStep(1);
}
 
function goToStep(n){
  currentWizardStep=n;
  document.querySelectorAll(".wizard-step-content").forEach(c=>c.style.display="none");
  document.getElementById(`wizard-step-${n}`).style.display="block";
  const nodes = document.querySelectorAll(".wizard-step-node");
  nodes.forEach((node,i)=>{
    node.className="wizard-step-node";
    if(i+1===n) node.classList.add("active-node");
    else if(i+1<n) node.classList.add("completed-node");
  });
  document.getElementById("wizard-line").style.width=`${((n-1)/(nodes.length-1))*100}%`;
  document.getElementById("wizard-step-indicator").innerText=`STEP ${n} OF ${nodes.length}`;
}
 
function nextWizardStep(){
  if(currentWizardStep===1){
    const vals = ["reg-card-number","reg-first-name","reg-last-name","reg-dob","reg-phone"];
    if(vals.some(id=>!document.getElementById(id).value.trim())){
      showToast("Please fill all mandatory fields.","error"); return;
    }
  }
  goToStep(currentWizardStep+1);
}
function prevWizardStep(){ goToStep(currentWizardStep-1); }
 
function calculateAgeFromDOB(){
  const dob = document.getElementById("reg-dob").value;
  if(!dob) return;
  const b = new Date(dob), t = new Date();
  let age = t.getFullYear()-b.getFullYear();
  if(t.getMonth()<b.getMonth()||(t.getMonth()===b.getMonth()&&t.getDate()<b.getDate())) age--;
  document.getElementById("reg-age").value=age;
}
 
function calculateBMI(){
  const h=parseFloat(document.getElementById("reg-height").value);
  const w=parseFloat(document.getElementById("reg-weight").value);
  const bmiField=document.getElementById("reg-bmi");
  const badge=document.getElementById("reg-bmi-badge");
  if(!h||!w||h<=0||w<=0){ bmiField.value=""; badge.style.display="none"; return; }
  const bmi=(w/((h/100)**2)).toFixed(1);
  bmiField.value=bmi; badge.style.display="inline-flex";
  if(bmi<18.5){ badge.innerText="Underweight"; badge.className="badge badge-yellow"; }
  else if(bmi<25){ badge.innerText="Normal"; badge.className="badge badge-green"; }
  else if(bmi<30){ badge.innerText="Overweight"; badge.className="badge badge-yellow"; }
  else{ badge.innerText="Obese"; badge.className="badge badge-red"; }
}
 
document.addEventListener("change",function(e){
  if(e.target&&e.target.type==="checkbox"){
    const pl=e.target.closest(".checkbox-label");
    if(pl){ pl.classList.toggle("checked",e.target.checked); }
  }
});
 
function previewPatientPhoto(event){
  const reader=new FileReader();
  reader.onload=function(){
    const p=document.getElementById("photo-upload-preview");
    p.innerText=""; p.style.backgroundImage=`url('${reader.result}')`;
    registeredPhotoData=reader.result;
  };
  reader.readAsDataURL(event.target.files[0]);
}
 
function handleWizardFileUpload(event,category){
  const file=event.target.files[0]; if(!file) return;
  uploadedFilesStore[category].push(file.name);
  const listIds={"Blood Report":"blood-files-list","ECG Report":"ecg-files-list","2D ECHO Report":"echo-files-list","TMT Report":"tmt-files-list","CAG":"cag-files-list","Other Investigations":"others-files-list"};
  const list=document.getElementById(listIds[category]);
  const item=document.createElement("div");
  item.className="file-preview-item";
  item.innerHTML=`<div class="file-preview-details"><span class="file-preview-icon">📄</span><span class="file-preview-name">${file.name}</span></div><button type="button" class="file-preview-remove" onclick="removeWizardFile(this,'${category}','${file.name}')">×</button>`;
  list.appendChild(item);
  showToast(`${file.name} uploaded.`,"info");
}
 
function removeWizardFile(btn,cat,name){
  uploadedFilesStore[cat]=uploadedFilesStore[cat].filter(f=>f!==name);
  btn.closest(".file-preview-item").remove();
}
 
function getNextPatientId(){
  let maxNum = 0;
  DB_PATIENTS.forEach(p=>{
    const m = (p.id||"").match(/\d+/);
    if(m){ const n=parseInt(m[0],10); if(n>maxNum) maxNum=n; }
  });
  return "PT-"+String(maxNum+1).padStart(3,"0");
}
 
// ===== SUBMIT NEW PATIENT WITH AUTO EMAIL =====
document.getElementById("wizard-patient-form").addEventListener("submit", function(e){
  e.preventDefault();
  const g = id => document.getElementById(id);
  const newPatient = {
    id: getNextPatientId(),
    cardType: g("reg-card-type").value, 
    cardNumber: g("reg-card-number").value.trim(),
    firstName: g("reg-first-name").value.trim(), 
    lastName: g("reg-last-name").value.trim(),
    dob: g("reg-dob").value, 
    age: parseInt(g("reg-age").value),
    gender: g("reg-gender").value, 
    photo: registeredPhotoData,
    occupation: g("reg-occupation").value.trim(), 
    address: g("reg-address").value.trim(),
    hometown: g("reg-hometown").value.trim(), 
    mandal: g("reg-mandal").value.trim(),
    district: g("reg-district").value.trim(), 
    state: g("reg-state").value.trim(),
    pincode: g("reg-pincode").value.trim(), 
    country: g("reg-country").value.trim(),
    phone: g("reg-phone").value.trim(), 
    email: g("reg-email").value.trim(),
    emName: g("reg-em-name").value.trim(), 
    emRelation: g("reg-em-relation").value.trim(),
    emPhone: g("reg-em-phone").value.trim(), 
    emEmail: g("reg-em-email").value.trim(),
    refBy: g("reg-ref-by").value, 
    refReason: g("reg-ref-reason").value.trim(),
    consultType: g("reg-consult-type").value,
    testsRequired: {ecg: g("test-ecg").checked, echo: g("test-2de").checked, tmt: g("test-tmt").checked},
    cvsSymptoms: g("reg-symptoms-cvs").value.trim(), 
    swelling: g("reg-swelling").value.trim(),
    otherSymptoms: g("reg-symptoms-other").value.trim(),
    presentIllness: g("reg-present-illness").value.trim(),
    pastIllness: g("reg-past-illness").value.trim(),
    riskFactors: {dm: g("risk-dm").checked, htn: g("risk-htn").checked, smoke: g("risk-smoke").checked, ihd: g("risk-ihd").checked, others: g("risk-others").checked},
    pastInvestigations: g("reg-past-investigations").value.trim(),
    hospitalization: g("reg-hospitalization").value.trim(),
    surgeryHistory: g("reg-surgery-history").value.trim(),
    drugHistory: g("reg-drug-history").value.trim(),
    appetite: g("reg-appetite").value.trim(), 
    stools: g("reg-stools").value.trim(),
    urination: g("reg-urination").value.trim(),
    familyHistory: g("reg-family-history").value.trim(),
    menstrual: g("reg-menstrual").value.trim(),
    immunization: g("reg-immunization").value.trim(),
    socioeconomic: g("reg-socioeconomic").value.trim(),
    height: parseFloat(g("reg-height").value) || "", 
    weight: parseFloat(g("reg-weight").value) || "",
    bmi: g("reg-bmi").value, 
    pulse: g("reg-pulse").value.trim(),
    rhythm: g("reg-rhythm").value.trim(), 
    peripheral: g("reg-peripheral").value.trim(),
    supineBP: g("reg-supine-bp").value.trim(), 
    sittingBP: g("reg-sitting-bp").value.trim(),
    standingBP: g("reg-standing-bp").value.trim(),
    cvsExam: g("reg-cvs-exam").value.trim(), 
    otherExam: g("reg-other-exam").value.trim(),
    files: Object.values(uploadedFilesStore).flat(),
    management: g("reg-management").value.trim(),
    treatment: g("reg-treatment").value.trim(), 
    conclusion: g("reg-conclusion").value.trim(),
    lastEdited: null, 
    editHistory: []
  };
  
  DB_PATIENTS.push(newPatient);
  savePatients();
  showToast(`✅ Patient ${newPatient.firstName} ${newPatient.lastName} saved!`, "success");
  
  // Send confirmation email
  sendPatientEmail(newPatient);
  
  setTimeout(() => exportAllPatientsToExcel(), 1000);
  switchView("old-patient");
});
 
// ===== PATIENT DIRECTORY =====
function renderPatientDirectory(){
  const body=document.getElementById("patient-table-body");
  body.innerHTML="";
 
  const serialMap = {};
  DB_PATIENTS.forEach((p,i)=>{ serialMap[p.id] = i+1; });
 
  let filtered=[...DB_PATIENTS];
  if(currentPatientFilter==="DM") filtered=filtered.filter(p=>p.riskFactors.dm);
  if(currentPatientFilter==="HTN") filtered=filtered.filter(p=>p.riskFactors.htn);
  if(currentPatientFilter==="IHD") filtered=filtered.filter(p=>p.riskFactors.ihd);
  const s=document.getElementById("patient-search").value.toLowerCase().trim();
  if(s) filtered=filtered.filter(p=>
    p.firstName.toLowerCase().includes(s)||p.lastName.toLowerCase().includes(s)||
    p.id.toLowerCase().includes(s)||p.phone.includes(s)||p.hometown.toLowerCase().includes(s)
  );
 
  filtered.sort((a,b)=>serialMap[a.id]-serialMap[b.id]);
 
  document.getElementById("directory-total-badge").innerText=`${filtered.length} PATIENTS`;
  if(!filtered.length){
    body.innerHTML=`<tr><td colspan="9" style="text-align:center;padding:40px;color:var(--text-secondary);">No patients found.</td></tr>`;
    return;
  }
  filtered.forEach(p=>{
    let rb="";
    if(p.riskFactors.dm) rb+=`<span class="badge badge-blue" style="font-size:8px;margin-right:3px;">DM</span>`;
    if(p.riskFactors.htn) rb+=`<span class="badge badge-red" style="font-size:8px;margin-right:3px;">HTN</span>`;
    if(p.riskFactors.smoke) rb+=`<span class="badge badge-yellow" style="font-size:8px;margin-right:3px;">SMOKE</span>`;
    if(p.riskFactors.ihd) rb+=`<span class="badge badge-green" style="font-size:8px;">IHD</span>`;
    if(!rb) rb=`<span style="color:var(--text-muted);font-size:11px;">None</span>`;
    const lastUpdatedCell = p.lastEdited
      ? `<span class="badge badge-yellow" style="font-size:9px;" title="${(p.editHistory&&p.editHistory.length)||0} edit(s) recorded">✏️ ${formatDateTime(p.lastEdited)}</span>`
      : `<span style="color:var(--text-muted);font-size:11px;">Not edited</span>`;
    const tr=document.createElement("tr");
    tr.innerHTML=`
      <td><strong>${serialMap[p.id]}</strong></td>
      <td><strong>${p.id}</strong></td>
      <td><strong>${p.firstName} ${p.lastName}</strong><br><small style="color:var(--text-secondary);">${p.phone}</small></td>
      <td>${p.age} Yrs / ${p.gender}</td>
      <td>${p.hometown}, ${p.state}</td>
      <td><div>${rb}</div></td>
      <td><span class="badge badge-blue">${p.refBy}</span></td>
      <td>${lastUpdatedCell}</td>
      <td style="text-align:center;">
        <div style="display:flex;gap:6px;justify-content:center;">
          <button class="btn btn-secondary btn-icon" onclick="viewPatientCaseFile('${p.id}')" title="View Case File">📂</button>
          <button class="btn btn-primary btn-icon" onclick="openEditPatient('${p.id}')" title="Edit Patient" style="background-color:#8b5cf6;">✏️</button>
          <button class="btn btn-accent btn-icon" onclick="deletePatientRecord('${p.id}')" title="Delete">🗑️</button>
        </div>
      </td>`;
    body.appendChild(tr);
  });
}
 
function handleSearch(){ renderPatientDirectory(); }
function filterPatients(f){
  currentPatientFilter=f;
  document.querySelectorAll(".filter-chip").forEach(c=>c.classList.remove("active-chip"));
  const ids={all:"filter-all",DM:"filter-dm",HTN:"filter-htn",IHD:"filter-ihd"};
  document.getElementById(ids[f]).classList.add("active-chip");
  renderPatientDirectory();
}
 
function deletePatientRecord(id){
  if(!confirm(`Delete patient ${id}? This is permanent.`)) return;
  DB_PATIENTS=DB_PATIENTS.filter(p=>p.id!==id);
  DB_APPOINTMENTS=DB_APPOINTMENTS.filter(a=>a.patientId!==id);
  DB_INVOICES=DB_INVOICES.filter(i=>i.patientId!==id);
  savePatients(); saveAppointments(); saveInvoices();
  showToast("Patient record deleted.","info");
  renderPatientDirectory(); renderDashboard();
}
 
// ===== EDIT PATIENT =====
function openEditPatient(patientId){
  const p = DB_PATIENTS.find(x=>x.id===patientId);
  if(!p) return;
  editingPatientId = patientId;
  document.getElementById("edit-modal-title").innerText = `✏️ Edit Patient: ${p.firstName} ${p.lastName} (${p.id})`;
 
  const s = (id,val) => { const el=document.getElementById(id); if(el) el.value=val||""; };
  s("edit-first-name",p.firstName); s("edit-last-name",p.lastName);
  s("edit-dob",p.dob); s("edit-age",p.age); document.getElementById("edit-gender").value=p.gender||"Male";
  s("edit-phone",p.phone); s("edit-email",p.email); s("edit-occupation",p.occupation);
  s("edit-address",p.address); s("edit-hometown",p.hometown);
  s("edit-district",p.district); s("edit-state",p.state);
  document.getElementById("edit-ref-by").value=p.refBy||"Doctor";
  document.getElementById("edit-consult-type").value=p.consultType||"Routine";
  s("edit-ref-reason",p.refReason); s("edit-cvs-symptoms",p.cvsSymptoms);
  s("edit-present-illness",p.presentIllness); s("edit-drug-history",p.drugHistory);
  s("edit-family-history",p.familyHistory);
 
  document.getElementById("edit-risk-dm").checked = p.riskFactors.dm||false;
  document.getElementById("edit-risk-htn").checked = p.riskFactors.htn||false;
  document.getElementById("edit-risk-smoke").checked = p.riskFactors.smoke||false;
  document.getElementById("edit-risk-ihd").checked = p.riskFactors.ihd||false;
  document.getElementById("edit-risk-others").checked = p.riskFactors.others||false;
 
  document.querySelectorAll("#edit-patient-form .checkbox-label").forEach(lbl=>{
    const chk = lbl.querySelector("input[type=checkbox]");
    if(chk) lbl.classList.toggle("checked", chk.checked);
  });
 
  s("edit-height",p.height); s("edit-weight",p.weight); s("edit-bmi",p.bmi);
  s("edit-pulse",p.pulse); s("edit-supine-bp",p.supineBP); s("edit-sitting-bp",p.sittingBP);
  s("edit-cvs-exam",p.cvsExam); s("edit-management",p.management);
  s("edit-treatment",p.treatment); s("edit-conclusion",p.conclusion);
 
  document.getElementById("edit-patient-modal").style.display="flex";
}
 
function calculateEditAge(){
  const dob = document.getElementById("edit-dob").value; if(!dob) return;
  const b=new Date(dob), t=new Date();
  let age=t.getFullYear()-b.getFullYear();
  if(t.getMonth()<b.getMonth()||(t.getMonth()===b.getMonth()&&t.getDate()<b.getDate())) age--;
  document.getElementById("edit-age").value=age;
}
 
function calculateEditBMI(){
  const h=parseFloat(document.getElementById("edit-height").value);
  const w=parseFloat(document.getElementById("edit-weight").value);
  if(!h||!w||h<=0||w<=0){ document.getElementById("edit-bmi").value=""; return; }
  document.getElementById("edit-bmi").value=(w/((h/100)**2)).toFixed(1);
}
 
document.getElementById("edit-patient-form").addEventListener("submit",function(e){
  e.preventDefault();
  const idx=DB_PATIENTS.findIndex(x=>x.id===editingPatientId);
  if(idx===-1){ showToast("Patient not found!","error"); return; }
 
  const g=id=>document.getElementById(id);
  const newVals = {
    firstName:g("edit-first-name").value.trim(),
    lastName:g("edit-last-name").value.trim(),
    dob:g("edit-dob").value,
    age:parseInt(g("edit-age").value)||DB_PATIENTS[idx].age,
    gender:g("edit-gender").value,
    phone:g("edit-phone").value.trim(),
    email:g("edit-email").value.trim(),
    occupation:g("edit-occupation").value.trim(),
    address:g("edit-address").value.trim(),
    hometown:g("edit-hometown").value.trim(),
    district:g("edit-district").value.trim(),
    state:g("edit-state").value.trim(),
    refBy:g("edit-ref-by").value,
    consultType:g("edit-consult-type").value,
    refReason:g("edit-ref-reason").value.trim(),
    cvsSymptoms:g("edit-cvs-symptoms").value.trim(),
    presentIllness:g("edit-present-illness").value.trim(),
    riskFactors:{
      dm:g("edit-risk-dm").checked,
      htn:g("edit-risk-htn").checked,
      smoke:g("edit-risk-smoke").checked,
      ihd:g("edit-risk-ihd").checked,
      others:g("edit-risk-others").checked
    },
    drugHistory:g("edit-drug-history").value.trim(),
    familyHistory:g("edit-family-history").value.trim(),
    height:parseFloat(g("edit-height").value)||DB_PATIENTS[idx].height,
    weight:parseFloat(g("edit-weight").value)||DB_PATIENTS[idx].weight,
    bmi:g("edit-bmi").value,
    pulse:g("edit-pulse").value.trim(),
    supineBP:g("edit-supine-bp").value.trim(),
    sittingBP:g("edit-sitting-bp").value.trim(),
    cvsExam:g("edit-cvs-exam").value.trim(),
    management:g("edit-management").value.trim(),
    treatment:g("edit-treatment").value.trim(),
    conclusion:g("edit-conclusion").value.trim()
  };
 
  const oldPatientSnapshot = DB_PATIENTS[idx];
  const changeList = buildEditDiff(oldPatientSnapshot, newVals);
  const nowIso = new Date().toISOString();
 
  DB_PATIENTS[idx] = {
    ...oldPatientSnapshot,
    ...newVals,
    lastEdited: nowIso,
    editHistory: [
      ...(oldPatientSnapshot.editHistory||[]),
      ...(changeList.length ? [{date:nowIso, changes:changeList}] : [])
    ]
  };
 
  savePatients();
  closeModal("edit-patient-modal");
  renderPatientDirectory();
  renderDashboard();
  const changeMsg = changeList.length
    ? `${changeList.length} field${changeList.length>1?'s':''} updated.`
    : "No field changes detected.";
  showToast(`Patient ${DB_PATIENTS[idx].firstName} updated on ${formatDateTime(nowIso)}. ${changeMsg}`,"success");
  setTimeout(()=>exportAllPatientsToExcel(), 800);
});
 
// ===== VIEW CASE FILE =====
function viewPatientCaseFile(patientId){
  const p=DB_PATIENTS.find(x=>x.id===patientId); if(!p) return;
  document.getElementById("modal-title-name").innerText=`Clinical File: ${p.firstName} ${p.lastName} (${p.id})`;
  const init=`${p.firstName.charAt(0)}${p.lastName.charAt(0)}`.toUpperCase();
  const photoStyle=p.photo?`style="background-image:url('${p.photo}')"` :'';
  const tests=[p.testsRequired?.ecg?"ECG":"",p.testsRequired?.echo?"2D Echo":"",p.testsRequired?.tmt?"TMT":""].filter(Boolean).join(", ")||"None";
  const risks=[p.riskFactors?.dm?"DM":"",p.riskFactors?.htn?"HTN":"",p.riskFactors?.smoke?"Smoking":"",p.riskFactors?.ihd?"Family IHD":""].filter(Boolean).join(", ")||"None";
  const billed=DB_INVOICES.filter(i=>i.patientId===patientId).reduce((a,b)=>a+b.amount,0);
  const lastEditedBadge = p.lastEdited
    ? `<span class="badge badge-yellow" style="font-size:11px;">✏️ Last Updated: ${formatDateTime(p.lastEdited)}</span>`
    : `<span class="badge" style="font-size:11px;background-color:var(--bg-page);color:var(--text-muted);">✏️ Never Edited</span>`;
  document.getElementById("modal-body-content").innerHTML=`
    <div class="clinical-header">
      <div class="clinical-avatar" ${photoStyle}>${p.photo?'':init}</div>
      <div class="clinical-meta-info">
        <h2>${p.firstName} ${p.lastName}</h2>
        <div class="clinical-meta-details">
          <span><strong>Sex/Age:</strong> ${p.gender}, ${p.age} Yrs</span>
          <span><strong>Card:</strong> ${p.cardType} (${p.cardNumber})</span>
          <span><strong>Phone:</strong> ${p.phone}</span>
        </div>
      </div>
      <div style="display:flex;flex-direction:column;gap:6px;align-items:flex-end;">
        <span class="badge badge-blue" style="font-size:12px;">₹${billed.toLocaleString('en-IN')} Billed</span>
        ${lastEditedBadge}
      </div>
    </div>
    <div class="form-section">
      <h3 style="margin-bottom:12px;color:var(--primary);">1. Contact & Demographics</h3>
      <div class="clinical-grid">
        <div class="clinical-item"><div class="clinical-label">Address</div><div class="clinical-value">${p.address||'N/A'}</div></div>
        <div class="clinical-item"><div class="clinical-label">City/Town</div><div class="clinical-value">${p.hometown||'N/A'}</div></div>
        <div class="clinical-item"><div class="clinical-label">District</div><div class="clinical-value">${p.district||'N/A'}</div></div>
      </div>
    </div>
    <div class="form-section">
      <h3 style="margin-bottom:12px;color:var(--primary);">2. Medical History</h3>
      <div class="clinical-grid">
        <div class="clinical-item"><div class="clinical-label">Referred By</div><div class="clinical-value">${p.refBy}</div></div>
        <div class="clinical-item"><div class="clinical-label">Consultation</div><div class="clinical-value">${p.consultType}</div></div>
        <div class="clinical-item"><div class="clinical-label">Tests</div><div class="clinical-value">${tests}</div></div>
      </div>
      <div style="margin-top:12px;display:flex;flex-direction:column;gap:8px;">
        <div><strong>CVS Symptoms:</strong> <span>${p.cvsSymptoms||'N/A'}</span></div>
        <div><strong>Risk Factors:</strong> <span>${risks}</span></div>
        <div><strong>Present Illness:</strong> <span>${p.presentIllness||'N/A'}</span></div>
        <div><strong>Drug History:</strong> <span>${p.drugHistory||'N/A'}</span></div>
      </div>
    </div>
    <div class="form-section">
      <h3 style="margin-bottom:12px;color:var(--primary);">3. Vitals</h3>
      <div class="clinical-grid">
        <div class="clinical-item"><div class="clinical-label">BP (Supine/Sitting)</div><div class="clinical-value">${p.supineBP||'N/A'} / ${p.sittingBP||'N/A'}</div></div>
        <div class="clinical-item"><div class="clinical-label">Pulse</div><div class="clinical-value">${p.pulse||'N/A'}</div></div>
        <div class="clinical-item"><div class="clinical-label">BMI</div><div class="clinical-value">HT: ${p.height}cm • WT: ${p.weight}kg • BMI: ${p.bmi||'N/A'}</div></div>
      </div>
    </div>
    <div class="form-section" style="border-left-color:var(--success);">
      <h3 style="margin-bottom:12px;color:var(--success);">4. Conclusion & Treatment</h3>
      <div style="display:flex;flex-direction:column;gap:12px;">
        <div><strong style="color:var(--success);">Conclusion:</strong><p style="margin-top:4px;background:var(--bg-page);padding:10px;border-radius:5px;border:1px solid var(--border-color);">${p.conclusion||'Pending.'}</p></div>
        <div><strong>Treatment:</strong><pre style="font-family:inherit;white-space:pre-wrap;margin-top:4px;background:var(--bg-page);padding:10px;border-radius:5px;border:1px solid var(--border-color);">${p.treatment||'Not written yet.'}</pre></div>
        <div><strong>Management Plan:</strong><p style="margin-top:4px;">${p.management||'N/A'}</p></div>
      </div>
    </div>
    <div class="form-section" style="border-left-color:var(--warning);">
      <h3 style="margin-bottom:12px;color:var(--warning);">5. Edit History & Audit Trail</h3>
      ${buildEditHistoryHtml(p)}
    </div>`;
  document.getElementById("patient-details-modal").style.display="flex";
}
 
function closeModal(id){ document.getElementById(id).style.display="none"; }
 
function printClinicalFile(){
  const content=document.getElementById("modal-body-content").innerHTML;
  const style=`<style>body{font-family:Arial,sans-serif;padding:40px;color:#000;}.form-section{border:1px solid #ccc;padding:15px;margin-bottom:20px;border-radius:5px;}.clinical-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;margin-bottom:15px;}.clinical-item{border:1px solid #ddd;padding:10px;}.clinical-label{font-size:9px;font-weight:bold;text-transform:uppercase;color:#555;}.clinical-value{font-size:12px;margin-top:2px;}pre{white-space:pre-wrap;font-family:inherit;}</style>`;
  const w=window.open("","_blank");
  w.document.write(`<html><head><title>Clinical File</title>${style}</head><body>${content}</body></html>`);
  w.document.close(); w.focus(); setTimeout(()=>{ w.print(); w.close(); },250);
}
 
window.addEventListener("click",function(e){
  ["patient-details-modal","edit-patient-modal"].forEach(id=>{
    if(e.target===document.getElementById(id)) closeModal(id);
  });
});
 
// ===== APPOINTMENTS =====
function populatePatientDropdowns(){
  ["book-patient-select","email-patient-select","bill-patient-select"].forEach(id=>{
    const sel=document.getElementById(id);
    sel.innerHTML=`<option value="">-- Select Patient --</option>`;
    DB_PATIENTS.forEach(p=>{
      const o=document.createElement("option");
      o.value=p.id; o.innerText=`${p.firstName} ${p.lastName} (${p.id})`;
      sel.appendChild(o);
    });
  });
}
 
function generateTimeSlots(){
  const c=document.getElementById("booking-slots-container"); c.innerHTML="";
  const d=document.getElementById("book-date").value;
  if(!d){ c.innerHTML="<small style='color:var(--text-muted);'>Please select a date first.</small>"; return; }
  const times=["09:00 AM","09:30 AM","10:00 AM","10:30 AM","11:00 AM","11:30 AM","02:00 PM","02:30 PM","03:00 PM","03:30 PM","04:00 PM","04:30 PM"];
  const booked=DB_APPOINTMENTS.filter(a=>a.date===d&&a.status!=="Cancelled").map(a=>a.time);
  times.forEach(t=>{
    const btn=document.createElement("button");
    btn.type="button"; btn.className="slot-btn"; btn.innerText=t;
    if(booked.includes(t)){
      btn.classList.add("disabled-slot"); btn.title="Already booked";
    } else {
      btn.onclick=()=>{
        document.querySelectorAll(".slot-btn").forEach(b=>b.classList.remove("active-slot"));
        btn.classList.add("active-slot");
        document.getElementById("selected-time-slot").value=t;
      };
    }
    c.appendChild(btn);
  });
}
 
document.getElementById("appointment-booking-form").addEventListener("submit",function(e){
  e.preventDefault();
  const pid=document.getElementById("book-patient-select").value;
  const date=document.getElementById("book-date").value;
  const time=document.getElementById("selected-time-slot").value;
  if(!pid||!date||!time){ showToast("Please select patient, date and time slot.","error"); return; }
  DB_APPOINTMENTS.push({id:"AP-"+Math.floor(100+Math.random()*900),patientId:pid,date,time,consultType:document.getElementById("book-consultation").value,notes:document.getElementById("book-notes").value.trim(),status:"Scheduled"});
  saveAppointments();
  showToast(`Appointment booked at ${time}!`,"success");
  switchView("appointment-view");
});
 
function renderAppointmentsPlanner(){
  const c=document.getElementById("appointment-planner-list"); c.innerHTML="";
  const today=new Date().toISOString().split('T')[0];
  let filtered=[...DB_APPOINTMENTS];
  if(currentApptFilter==="today") filtered=filtered.filter(a=>a.date===today);
  else if(currentApptFilter==="upcoming") filtered=filtered.filter(a=>a.date>today);
  else if(currentApptFilter==="completed") filtered=filtered.filter(a=>a.status==="Completed");
  filtered.sort((a,b)=>a.date.localeCompare(b.date)||a.time.localeCompare(b.time));
  if(!filtered.length){ c.innerHTML=`<div style="text-align:center;padding:40px;color:var(--text-secondary);">No appointments found.</div>`; return; }
  filtered.forEach(a=>{
    const p=DB_PATIENTS.find(x=>x.id===a.patientId); if(!p) return;
    const sc={Completed:"badge-green",Cancelled:"badge-red","Checked-In":"badge-yellow"};
    const bc=sc[a.status]||"badge-blue";
    let btns="";
    if(a.status==="Scheduled") btns=`<button class="btn btn-secondary" style="font-size:11px;padding:6px 10px;" onclick="updateApptStatus('${a.id}','Checked-In')">Check-In</button><button class="btn btn-accent" style="font-size:11px;padding:6px 10px;" onclick="updateApptStatus('${a.id}','Cancelled')">Cancel</button>`;
    else if(a.status==="Checked-In") btns=`<button class="btn btn-success" style="font-size:11px;padding:6px 10px;" onclick="updateApptStatus('${a.id}','Completed')">Complete</button>`;
    const div=document.createElement("div");
    div.className=`appointment-card ${a.consultType==="Emergency"?"urgent":""}`;
    div.innerHTML=`<div style="display:flex;align-items:center;gap:20px;"><div class="ap-time">${a.time}<br><small style="color:var(--text-muted);">${a.date}</small></div><div class="ap-details"><h4>${p.firstName} ${p.lastName} (${p.id})</h4><p>${a.notes||'—'} • <span class="badge ${bc}" style="font-size:8px;">${a.status}</span></p></div></div><div class="ap-actions">${btns}<button class="btn btn-secondary btn-icon" onclick="viewPatientCaseFile('${p.id}')">📂</button></div>`;
    c.appendChild(div);
  });
}
 
function updateApptStatus(id,status){
  const a=DB_APPOINTMENTS.find(x=>x.id===id); if(!a) return;
  a.status=status; saveAppointments();
  showToast(`Status updated to ${status}.`,"success");
  renderAppointmentsPlanner(); renderDashboard();
}
 
function filterAppointments(f){
  currentApptFilter=f;
  ["today","upcoming","completed"].forEach(x=>{ document.getElementById(`appt-filter-${x}`).classList.remove("active-chip"); });
  document.getElementById(`appt-filter-${f}`).classList.add("active-chip");
  renderAppointmentsPlanner();
}
 
// ===== CALCULATORS =====
function runRiskCalculator(){
  const age=parseInt(document.getElementById("calc-age").value);
  const sbp=parseInt(document.getElementById("calc-sbp").value);
  const chol=parseInt(document.getElementById("calc-chol").value);
  const hdl=parseInt(document.getElementById("calc-hdl").value);
  const dm=document.getElementById("calc-dm").value==="yes";
  const smoke=document.getElementById("calc-smoke").value==="yes";
  document.getElementById("calc-age-val").innerText=age;
  document.getElementById("calc-sbp-val").innerText=sbp;
  document.getElementById("calc-chol-val").innerText=chol;
  document.getElementById("calc-hdl-val").innerText=hdl;
  let r=(age-20)*0.15+(sbp-90)*0.12+(chol-100)*0.08-(hdl-20)*0.1+(dm?6:0)+(smoke?8:0);
  r=Math.max(1,r); const fr=Math.min(r,40).toFixed(1);
  document.getElementById("calc-risk-result").innerText=`${fr}%`;
  const box=document.getElementById("calc-risk-box");
  const stat=document.getElementById("calc-risk-status");
  if(fr<10){ stat.innerText="LOW RISK (<10%)"; box.style.backgroundColor="rgba(16,185,129,0.1)"; box.style.border="1px solid rgba(16,185,129,0.2)"; stat.style.color="var(--success)"; }
  else if(fr<20){ stat.innerText="MODERATE RISK (10-20%)"; box.style.backgroundColor="rgba(245,158,11,0.1)"; box.style.border="1px solid rgba(245,158,11,0.2)"; stat.style.color="var(--warning)"; }
  else{ stat.innerText="HIGH RISK (>20%)"; box.style.backgroundColor="rgba(239,68,68,0.1)"; box.style.border="1px solid rgba(239,68,68,0.2)"; stat.style.color="var(--accent)"; }
}
 
function calculateTHR(){
  const age=parseInt(document.getElementById("thr-age").value)||50;
  const max=220-age;
  document.getElementById("thr-max").innerText=`${max} bpm`;
  document.getElementById("thr-zone").innerText=`${Math.round(max*0.7)} - ${Math.round(max*0.85)} bpm`;
}
 
// ===== EMAIL TEMPLATES =====
const emailTemplates={
  confirmation:{subject:"Appointment Confirmation - Srinivasa Heart Centre",body:"Dear [NAME],\n\nYour cardiological consultation at Srinivasa Heart Centre is confirmed.\n\nDate: [DATE]\nTime: [TIME]\nConsultant: Dr. Srinivas Ramaka, MD, DM.\n\nPlease arrive 10 minutes early.\n\nRegards,\nSrinivasa Heart Centre"},
  reports:{subject:"Diagnostic Reports Ready",body:"Dear [NAME],\n\nYour diagnostic reports are ready.\n\nFinal Recommendation: [CONCLUSION]\n\nPlease consult the clinic for details.\n\nRegards,\nSrinivasa Heart Centre"},
  guidelines:{subject:"Cardiac Health Recommendations",body:"Dear [NAME],\n\nDr. Srinivas Ramaka recommends:\n\n1. Low-sodium, high-fiber diet.\n2. 30 minutes walking daily.\n3. Monitor blood pressure regularly.\n4. Take medications as prescribed.\n\nBest health,\nDr. Srinivas Ramaka"},
  custom:{subject:"Clinic Notification",body:"Dear [NAME],\n\n\n\nRegards,\nSrinivasa Heart Centre"}
};
 
function autoFillEmailRecipient(){
  const id=document.getElementById("email-patient-select").value;
  const p=DB_PATIENTS.find(x=>x.id===id);
  document.getElementById("email-to").value=p?p.email||"":"";
  applyEmailTemplate();
}
 
function applyEmailTemplate(){
  const type=document.getElementById("email-template-select").value;
  const id=document.getElementById("email-patient-select").value;
  const p=DB_PATIENTS.find(x=>x.id===id)||DB_PATIENTS[0];
  const tmpl=emailTemplates[type];
  let body=tmpl.body, subj=tmpl.subject;
  if(p){
    body=body.replace("[NAME]",`${p.firstName} ${p.lastName}`);
    const appt=DB_APPOINTMENTS.find(a=>a.patientId===p.id);
    body=body.replace("[DATE]",appt?appt.date:new Date().toISOString().split('T')[0]);
    body=body.replace("[TIME]",appt?appt.time:"09:30 AM");
    body=body.replace("[CONCLUSION]",p.conclusion||"Stable vitals.");
  }
  document.getElementById("email-subject").value=subj;
  document.getElementById("email-body").value=body;
  syncEmailPreview();
}
 
function syncEmailPreview(){
  document.getElementById("preview-subj").innerText=document.getElementById("email-subject").value;
  document.getElementById("preview-body-content").innerHTML=document.getElementById("email-body").value.replace(/\n/g,"<br>");
}
 
// ===== SEND EMAIL FROM EMAIL PORTAL =====
// FIX: trims the recipient field, validates it BEFORE calling EmailJS (so a
// blank/space-only field never reaches the API), and sends the same address
// under several common variable names (to_email / email / recipient /
// reply_to) so it matches whichever one your EmailJS "To Email" setting uses.
document.getElementById("email-compose-form").addEventListener("submit", async function(e){
  e.preventDefault();
  
  const toEmail = document.getElementById("email-to").value.trim();
  const subject = document.getElementById("email-subject").value.trim();
  const body = document.getElementById("email-body").value.trim();
  const patientId = document.getElementById("email-patient-select").value;
  const patient = DB_PATIENTS.find(p => p.id === patientId);
  const patientName = patient ? `${patient.firstName} ${patient.lastName}` : "Patient";
  
  // Validate BEFORE calling EmailJS so we never send a blank recipient.
  if(!toEmail){
    showToast("⚠️ Please enter a recipient email address — the 'To' field is empty.", "error");
    document.getElementById("email-to").focus();
    return;
  }
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if(!emailPattern.test(toEmail)){
    showToast("⚠️ That doesn't look like a valid email address.", "error");
    document.getElementById("email-to").focus();
    return;
  }
  if(!subject || !body){
    showToast("Please fill the subject and message body.", "error");
    return;
  }
  
  const sendBtn = document.getElementById("send-email-btn");
  sendBtn.disabled = true;
  sendBtn.innerHTML = "⏳ Sending...";
  
  try {
    // Send the address under multiple common variable names so it matches
    // whichever one is bound to "To Email" in your EmailJS template/service.
    const templateParams = {
      to_email: toEmail,
      email: toEmail,
      recipient: toEmail,
      reply_to: toEmail,
      to_name: patientName,
      subject: subject,
      message: body,
      from_name: "Dr. Srinivas Ramaka",
      clinic_name: "Srinivasa Heart Centre"
    };
    
    console.log("📧 Sending email with params:", templateParams);
    
    const response = await emailjs.send(
      EMAILJS_CONFIG.SERVICE_ID,
      EMAILJS_CONFIG.TEMPLATE_ID,
      templateParams
    );
    
    console.log("✅ Email sent successfully:", response);
    showToast(`✅ Email successfully sent to ${toEmail}!`, "success");
    
    document.getElementById("email-compose-form").reset();
    applyEmailTemplate();
    populatePatientDropdowns();
    
  } catch(error) {
    console.error("❌ Email sending failed:", error);
    const errText = (error && (error.text || error.message)) || "Unknown error";
    if(/recipient/i.test(errText) && /empty/i.test(errText)){
      // This specific EmailJS error means the "To Email" field on your
      // EmailJS Service/Template settings page is not bound to a variable
      // we're sending (or is left blank) — it is a dashboard config issue,
      // not something the page's JS value can fix on its own.
      showToast(`❌ EmailJS says the recipient is empty. Open your EmailJS template → "To Email" and set it to {{to_email}}, then try again.`, "error");
    } else {
      showToast(`❌ Failed to send email: ${errText}`, "error");
    }
  } finally {
    sendBtn.disabled = false;
    sendBtn.innerHTML = "📧 Dispatch Patient Email";
  }
});
 
// ===== BILLING =====
function renderBillingLedger(){
  const total=DB_INVOICES.reduce((a,b)=>a+b.amount,0);
  const paid=DB_INVOICES.filter(i=>i.status==="Paid").reduce((a,b)=>a+b.amount,0);
  const pend=total-paid;
  document.getElementById("billing-stat-total").innerText=`₹${total.toLocaleString('en-IN')}`;
  document.getElementById("billing-stat-paid").innerText=`₹${paid.toLocaleString('en-IN')}`;
  document.getElementById("billing-stat-pending").innerText=`₹${pend.toLocaleString('en-IN')}`;
  const body=document.getElementById("billing-ledger-body"); body.innerHTML="";
  if(!DB_INVOICES.length){ body.innerHTML="<tr><td colspan='6' style='text-align:center;'>No invoices yet.</td></tr>"; return; }
  DB_INVOICES.forEach(inv=>{
    const p=DB_PATIENTS.find(x=>x.id===inv.patientId);
    const name=p?`${p.firstName} ${p.lastName}`:"Unknown";
    const method=inv.paymentMethod||"Cash";
    const tr=document.createElement("tr");
    tr.innerHTML=`<td><strong>${name}</strong><br><small style="color:var(--text-muted);">${inv.id} • ${inv.date}</small></td><td style="max-width:160px;"><small>${inv.items.join(", ")}</small></td><td><strong>₹${inv.amount.toLocaleString('en-IN')}</strong></td><td><span class="badge badge-blue" style="font-size:9px;">${method}</span></td><td><span class="badge ${inv.status==='Paid'?'badge-green':'badge-red'}">${inv.status}</span></td><td>${inv.status==="Pending"?`<button class="btn btn-secondary" style="font-size:10px;padding:4px 8px;" onclick="collectPayment('${inv.id}')">Collect</button>`:`<small style="color:var(--text-muted);">Settled</small>`}</td>`;
    body.appendChild(tr);
  });
}
 
function collectPayment(id){
  const inv=DB_INVOICES.find(x=>x.id===id); if(!inv) return;
  inv.status="Paid"; saveInvoices();
  showToast(`Payment collected for ${id}!`,"success");
  renderBillingLedger(); renderDashboard();
}
 
document.querySelectorAll(".bill-item-chk").forEach(c=>c.addEventListener("change",calcBillSum));
function calcBillSum(){
  const total=[...document.querySelectorAll(".bill-item-chk")].filter(c=>c.checked).reduce((a,c)=>a+parseInt(c.value),0);
  document.getElementById("bill-running-total").innerText=`₹${total.toLocaleString('en-IN')}`;
}
 
// ===== UPI / SCANNER PAYMENT QR =====
// Shows a scannable UPI QR code (PhonePe/Google Pay/Paytm all read standard
// UPI deep links) whenever a UPI payment method is selected, encoding the
// clinic's UPI ID and the current invoice total as the payable amount.
const CLINIC_UPI_ID = "srinivasaheartcentre@upi";
const CLINIC_UPI_NAME = "Srinivasa Heart Centre";

function togglePaymentQR(){
  const method = document.getElementById("bill-payment-method").value;
  const box = document.getElementById("upi-qr-box");
  if(method.startsWith("UPI")){
    box.style.display = "block";
    renderUpiQR();
  } else {
    box.style.display = "none";
  }
}

function renderUpiQR(){
  const wrapper = document.getElementById("upi-qr-canvas-wrapper");
  wrapper.innerHTML = "";
  const amount = [...document.querySelectorAll(".bill-item-chk")].filter(c=>c.checked).reduce((a,c)=>a+parseInt(c.value),0) || 0;
  const upiLink = `upi://pay?pa=${encodeURIComponent(CLINIC_UPI_ID)}&pn=${encodeURIComponent(CLINIC_UPI_NAME)}&am=${amount}&cu=INR`;
  if(typeof QRCode === "undefined"){
    wrapper.innerHTML = `<div style="font-size:11px;color:var(--text-muted);">QR library unavailable — share UPI ID manually.</div>`;
    return;
  }
  const canvas = document.createElement("canvas");
  wrapper.appendChild(canvas);
  QRCode.toCanvas(canvas, upiLink, { width: 160, margin: 1 }, function(err){
    if(err){
      wrapper.innerHTML = `<div style="font-size:11px;color:var(--text-muted);">Couldn't render QR — share UPI ID manually.</div>`;
    }
  });
}
// Re-render the QR amount live as services are toggled, if a UPI method is selected
document.querySelectorAll(".bill-item-chk").forEach(c=>c.addEventListener("change",()=>{
  if(document.getElementById("bill-payment-method").value.startsWith("UPI")) renderUpiQR();
}));

// ===== EMAIL INVOICE TO PATIENT =====
// Sends the generated invoice details to the patient's registered email via
// EmailJS, reusing the same multi-key recipient params as the rest of the
// app so it works regardless of how "To Email" is bound in the template.
function emailInvoiceToPatient(invoice, patient){
  const patientEmail = (patient && patient.email || "").trim();
  if(!patientEmail){
    showToast(`Invoice ${invoice.id} saved. Patient has no email on file, so it wasn't emailed.`, "info");
    return Promise.resolve("No email on file");
  }
  const patientName = `${patient.firstName} ${patient.lastName}`;
  const itemLines = invoice.items.join(", ");
  const message =
`Dear ${patientName},

Thank you for visiting Srinivasa Heart Centre. Please find your invoice details below.

Invoice ID: ${invoice.id}
Date: ${invoice.date}
Services: ${itemLines}
Total Amount: ₹${invoice.amount.toLocaleString('en-IN')}
Payment Method: ${invoice.paymentMethod}
Payment Status: ${invoice.status}

Thank you,
Dr. Srinivas Ramaka, MD, DM.
Srinivasa Heart Centre`;

  const templateParams = {
    to_email: patientEmail,
    email: patientEmail,
    recipient: patientEmail,
    reply_to: patientEmail,
    to_name: patientName,
    subject: `Invoice ${invoice.id} - Srinivasa Heart Centre`,
    message: message,
    from_name: "Dr. Srinivas Ramaka",
    clinic_name: "Srinivasa Heart Centre"
  };

  return emailjs.send(EMAILJS_CONFIG.SERVICE_ID, EMAILJS_CONFIG.TEMPLATE_ID, templateParams)
    .then(function(response){
      console.log("✅ Invoice email sent:", response);
      showToast(`✅ Invoice ${invoice.id} emailed to ${patientEmail}!`, "success");
      return response;
    })
    .catch(function(error){
      console.error("❌ Invoice email failed:", error);
      const errText = (error && (error.text || error.message)) || "Unknown error";
      showToast(`⚠️ Invoice ${invoice.id} saved, but email failed: ${errText}`, "error");
      return error;
    });
}
 
document.getElementById("billing-invoice-form").addEventListener("submit",function(e){
  e.preventDefault();
  const pid=document.getElementById("bill-patient-select").value;
  if(!pid){ showToast("Please select a patient.","error"); return; }
  const items=[],amounts=[];
  document.querySelectorAll(".bill-item-chk").forEach(c=>{ if(c.checked){ items.push(c.getAttribute("data-name")); amounts.push(parseInt(c.value)); } });
  if(!items.length){ showToast("Select at least one service.","error"); return; }
  const total=amounts.reduce((a,b)=>a+b,0);
  const paymentMethod=document.getElementById("bill-payment-method").value;
  const inv={
    id:"INV-"+Math.floor(100+Math.random()*900),
    patientId:pid,
    date:new Date().toISOString().split('T')[0],
    items,
    amount:total,
    status:document.getElementById("bill-invoice-status").value,
    paymentMethod:paymentMethod
  };
  DB_INVOICES.push(inv); saveInvoices();
  showToast(`Invoice ${inv.id} generated!`,"success");

  const patient = DB_PATIENTS.find(p=>p.id===pid);
  emailInvoiceToPatient(inv, patient);

  document.getElementById("billing-invoice-form").reset();
  document.querySelector(".bill-item-chk").checked=true;
  document.getElementById("upi-qr-box").style.display="none";
  calcBillSum(); renderBillingLedger(); renderDashboard();
});
 
// ===== THEME =====
function toggleTheme(){
  const t=document.documentElement.getAttribute("data-theme")==="dark"?"light":"dark";
  document.documentElement.setAttribute("data-theme",t);
  localStorage.setItem("SHC_THEME",t);
  if(document.getElementById("view-home").classList.contains("active-view")) initDashboardChart();
}
 
// ===== INIT =====
window.onload=function(){
  const theme=localStorage.getItem("SHC_THEME")||"light";
  document.documentElement.setAttribute("data-theme",theme);
  initClock();
  populatePatientDropdowns();
  renderDashboard();
  applyEmailTemplate();
  console.log("🚀 Clinic Portal loaded successfully!");
  console.log("📋 EmailJS Configuration:");
  console.log("   Public Key:", EMAILJS_CONFIG.PUBLIC_KEY);
  console.log("   Service ID:", EMAILJS_CONFIG.SERVICE_ID);
  console.log("   Template ID:", EMAILJS_CONFIG.TEMPLATE_ID);
};
</script>
</body>
</html>
