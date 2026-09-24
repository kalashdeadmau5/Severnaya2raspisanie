<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Расписание · МОУ "Северная СОШ №2"</title>
    <meta name="app-version" content="13.4">
    <meta name="app-name" content="Расписание">
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js"></script>
    <style>
        :root {
            --bg-page: #eef3f7;
            --bg-container: rgba(255, 255, 255, 0.94);
            --bg-container-border: rgba(255, 255, 255, 0.7);
            --text-primary: #0b2a4a;
            --text-secondary: #3d5a7a;
            --text-muted: #7a8fa8;
            --accent: #1a3a6b;
            --accent-hover: #0f2a4f;
            --accent-light: #e8eff8;
            --accent-lighter: #f0f7ff;
            --border: #d0ddee;
            --border-light: #dce5f0;
            --table-bg: #ffffff;
            --table-row-even: #fafcff;
            --table-row-hover: #f0f7ff;
            --th-bg: #1a3a6b;
            --th-bg-dark: #0f2a4f;
            --th-bg-darker: #0b2038;
            --teacher-bg: #e8eff8;
            --room-bg: #f5f8fc;
            --empty-color: #b0c0d0;
            --edit-hover: #fff8e7;
            --edit-outline: #d4a017;
            --success: #2a6b3a;
            --warning: #b8860b;
            --danger: #b00020;
            --substitute-bg: #fff3cd;
            --substitute-bg-even: #ffefb8;
            --substitute-bg-hover: #ffe89a;
            --substitute-border: #d4a017;
            --substitute-text: #7a5500;
            --substitute-badge-bg: #f0c040;
            --substitute-badge-text: #4a3500;
            --free-color: #2a6b3a;
            --busy-color: #b00020;
            --shadow-container: 0 12px 28px rgba(0, 20, 40, 0.12), 0 4px 8px rgba(0, 0, 0, 0.04);
            --shadow-table: 0 2px 8px rgba(0,0,0,0.03);
            --shadow-btn: 0 2px 6px rgba(0,0,0,0.08);
            --bg-word: rgba(30, 60, 120, 0.06);
            --bg-number: rgba(200, 140, 40, 0.07);
            --bg-day: rgba(30, 90, 140, 0.05);
            --modal-bg: #ffffff;
            --scrollbar-track: transparent;
            --scrollbar-thumb: rgba(26, 58, 107, 0.12);
            --scrollbar-thumb-hover: rgba(26, 58, 107, 0.3);
            --scrollbar-thumb-active: rgba(26, 58, 107, 0.45);
        }

        body.dark-theme {
            --bg-page: #0d1520;
            --bg-container: rgba(20, 30, 45, 0.95);
            --bg-container-border: rgba(255, 255, 255, 0.06);
            --text-primary: #e6eef8;
            --text-secondary: #a8bdd4;
            --text-muted: #6b7f96;
            --accent: #4a82c7;
            --accent-hover: #5e93d5;
            --accent-light: #1e2f4a;
            --accent-lighter: #1a2638;
            --border: #2a3c56;
            --border-light: #24344a;
            --table-bg: #16202f;
            --table-row-even: #1a2638;
            --table-row-hover: #223348;
            --th-bg: #1e3a5f;
            --th-bg-dark: #16304f;
            --th-bg-darker: #0f2440;
            --teacher-bg: #1e2f4a;
            --room-bg: #1a2638;
            --empty-color: #4a5f78;
            --edit-hover: #2a3a2f;
            --edit-outline: #d4a017;
            --success: #3d8f52;
            --warning: #d4a017;
            --danger: #e05070;
            --substitute-bg: #3a2f10;
            --substitute-bg-even: #42360f;
            --substitute-bg-hover: #4a3d12;
            --substitute-border: #d4a017;
            --substitute-text: #f0c040;
            --substitute-badge-bg: #d4a017;
            --substitute-badge-text: #1a1a1a;
            --free-color: #4ec06a;
            --busy-color: #e05070;
            --shadow-container: 0 12px 28px rgba(0, 0, 0, 0.5), 0 4px 8px rgba(0, 0, 0, 0.3);
            --shadow-table: 0 2px 8px rgba(0,0,0,0.25);
            --shadow-btn: 0 2px 6px rgba(0,0,0,0.3);
            --bg-word: rgba(120, 170, 240, 0.06);
            --bg-number: rgba(220, 170, 80, 0.07);
            --bg-day: rgba(120, 180, 240, 0.05);
            --modal-bg: #1a2638;
            --scrollbar-track: transparent;
            --scrollbar-thumb: rgba(168, 189, 212, 0.12);
            --scrollbar-thumb-hover: rgba(168, 189, 212, 0.3);
            --scrollbar-thumb-active: rgba(168, 189, 212, 0.45);
        }

        * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif; }

        html, body { height: 100%; overflow: hidden; }

        html, body, * {
            scrollbar-width: thin;
            scrollbar-color: var(--scrollbar-thumb) var(--scrollbar-track);
        }

        ::-webkit-scrollbar { width: 10px; height: 10px; background: transparent; }
        ::-webkit-scrollbar-track { background: transparent; border-radius: 10px; margin: 2px; }
        ::-webkit-scrollbar-thumb {
            background: var(--scrollbar-thumb);
            border-radius: 10px;
            border: 2px solid transparent;
            background-clip: padding-box;
            transition: background 0.2s ease;
        }
        ::-webkit-scrollbar-thumb:hover { background: var(--scrollbar-thumb-hover); background-clip: padding-box; }
        ::-webkit-scrollbar-thumb:active { background: var(--scrollbar-thumb-active); background-clip: padding-box; }
        ::-webkit-scrollbar-corner { background: transparent; }
        ::-webkit-resizer { background: transparent; }

        .console-body::-webkit-scrollbar { width: 8px; height: 8px; }
        .console-body::-webkit-scrollbar-thumb { background: var(--scrollbar-thumb); border-radius: 8px; border: 2px solid transparent; background-clip: padding-box; }
        .console-body::-webkit-scrollbar-thumb:hover { background: var(--scrollbar-thumb-hover); background-clip: padding-box; }

        .table-wrapper::-webkit-scrollbar { width: 10px; height: 10px; }
        .table-wrapper::-webkit-scrollbar-thumb { background: var(--scrollbar-thumb); border-radius: 10px; border: 2px solid transparent; background-clip: padding-box; }
        .table-wrapper::-webkit-scrollbar-thumb:hover { background: var(--scrollbar-thumb-hover); background-clip: padding-box; }

        body {
            background: var(--bg-page);
            position: relative;
            padding: 14px;
            display: flex;
            flex-direction: column;
            align-items: center;
            transition: background 0.3s ease;
            overflow: hidden;
        }

        .animated-bg { position: fixed; inset: 0; z-index: 0; pointer-events: none; overflow: hidden; }
        .animated-bg .word, .animated-bg .number, .animated-bg .day {
            position: absolute; top: 0; white-space: nowrap; user-select: none; animation: floatWord linear infinite;
        }
        .animated-bg .word   { color: var(--bg-word); font-weight: 700; }
        .animated-bg .number { color: var(--bg-number); font-weight: 800; }
        .animated-bg .day    { color: var(--bg-day); font-weight: 600; text-transform: uppercase; letter-spacing: 4px; }

        @keyframes floatWord {
            0%   { transform: translateY(110vh) rotate(0deg); opacity: 0; }
            10%  { opacity: 1; }
            90%  { opacity: 1; }
            100% { transform: translateY(-20vh) rotate(8deg); opacity: 0; }
        }

        .app-container {
            position: relative; z-index: 10; width: 100%; max-width: 1800px;
            height: 100%; max-height: calc(100vh - 28px);
            background: var(--bg-container); backdrop-filter: blur(8px);
            border-radius: 20px; box-shadow: var(--shadow-container);
            padding: 16px 20px 14px; border: 1px solid var(--bg-container-border);
            transition: background 0.3s ease, box-shadow 0.3s ease;
            display: flex; flex-direction: column; gap: 12px; overflow: hidden;
        }

        .school-header {
            display: flex; align-items: center; justify-content: space-between;
            padding-bottom: 10px; border-bottom: 2px solid var(--accent);
            flex-wrap: wrap; gap: 10px; flex-shrink: 0;
        }

        .school-title h1 { font-size: 1.35rem; font-weight: 700; color: var(--text-primary); letter-spacing: -0.3px; }
        .school-title p  { font-size: 0.82rem; color: var(--text-secondary); margin-top: 2px; font-weight: 500; }
        .school-title .subtitle { font-size: 0.76rem; color: var(--warning); font-weight: 600; margin-top: 1px; }

        .toolbar { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }

        .schedule-tabs {
            display: inline-flex; gap: 3px; padding: 3px;
            background: var(--accent-light); border: 1px solid var(--border);
            border-radius: 30px; margin-right: 8px; flex-shrink: 0;
        }

        .schedule-tab {
            display: inline-flex; align-items: center; gap: 5px;
            padding: 6px 14px; border-radius: 24px; border: none;
            background: transparent; color: var(--text-secondary);
            font-weight: 600; font-size: 0.76rem; font-family: inherit;
            cursor: pointer; transition: all 0.2s; white-space: nowrap;
        }

        .schedule-tab .tab-icon { font-size: 0.95rem; }

        .schedule-tab:hover:not(.active) { background: var(--accent-lighter); color: var(--accent); }

        .schedule-tab.active {
            background: var(--accent); color: white;
            box-shadow: 0 2px 6px rgba(26, 58, 107, 0.25);
            transform: translateY(-1px);
        }

        body.dark-theme .schedule-tab.active {
            background: var(--accent); box-shadow: 0 2px 6px rgba(74, 130, 199, 0.3);
        }

        .schedule-hint {
            display: flex; align-items: center; gap: 8px;
            padding: 8px 14px; background: var(--accent-lighter);
            border-left: 3px solid var(--accent); border-radius: 8px;
            font-size: 0.78rem; color: var(--text-secondary);
            flex-shrink: 0;
        }

        .schedule-hint .hint-icon { font-size: 1rem; }
        .schedule-hint strong { color: var(--accent); }

        body.mode-classes #substituteIndicator,
        body.mode-iup #substituteIndicator,
        body.mode-classes #adminIndicator,
        body.mode-iup #adminIndicator,
        body.mode-classes #logoutAdminBtn,
        body.mode-iup #logoutAdminBtn {
            display: none !important;
        }

        @media (max-width: 1100px) {
            .schedule-tab .tab-label { display: none; }
            .schedule-tab { padding: 8px 12px; }
            .schedule-tab .tab-icon { font-size: 1.1rem; }
        }

        .btn {
            padding: 7px 14px; border-radius: 30px; border: none; font-weight: 600; font-size: 0.8rem;
            cursor: pointer; transition: all 0.2s; display: inline-flex; align-items: center; gap: 6px;
            box-shadow: var(--shadow-btn); font-family: inherit; white-space: nowrap;
        }
        .btn-primary { background: var(--accent); color: white; }
        .btn-primary:hover { background: var(--accent-hover); transform: translateY(-1px); }
        .btn-outline { background: transparent; color: var(--accent); border: 1.5px solid var(--accent); }
        .btn-outline:hover { background: var(--accent-light); }
        .btn-success { background: var(--success); color: white; }
        .btn-success:hover { filter: brightness(1.1); }
        .btn-warning { background: var(--warning); color: white; }
        .btn-warning:hover { filter: brightness(1.1); }
        .btn-danger { background: var(--danger); color: white; }
        .btn-danger:hover { filter: brightness(1.1); }
        .btn-sm { padding: 4px 10px; font-size: 0.72rem; }
        .btn:disabled { opacity: 0.5; cursor: not-allowed; }

        .edit-indicator {
            display: none; align-items: center; gap: 6px; background: var(--accent-light);
            color: var(--success); border: 1.5px solid var(--success);
            padding: 6px 12px; border-radius: 30px; font-weight: 600; font-size: 0.76rem;
            animation: pulse 2s ease-in-out infinite;
        }
        .edit-indicator.visible { display: inline-flex; }
        .edit-indicator .dot { width: 8px; height: 8px; border-radius: 50%; background: var(--success); }
        @keyframes pulse {
            0%   { box-shadow: 0 0 0 0 rgba(42,107,58,0.4); }
            70%  { box-shadow: 0 0 0 10px rgba(42,107,58,0); }
            100% { box-shadow: 0 0 0 0 rgba(42,107,58,0); }
        }

        .substitute-indicator {
            display: none; align-items: center; gap: 6px; background: var(--substitute-bg);
            color: var(--substitute-text); border: 1.5px solid var(--substitute-border);
            padding: 6px 12px; border-radius: 30px; font-weight: 600; font-size: 0.76rem;
        }
        .substitute-indicator.visible { display: inline-flex; }

        .admin-mode-indicator {
            display: none; align-items: center; gap: 6px;
            background: linear-gradient(135deg, #1a3a6b, #2a5a9b);
            color: white; border: none; padding: 6px 12px; border-radius: 30px;
            font-weight: 600; font-size: 0.76rem;
            box-shadow: 0 2px 6px rgba(26, 58, 107, 0.3);
            animation: adminPulse 3s ease-in-out infinite;
        }
        .admin-mode-indicator.visible { display: inline-flex; }
        .admin-mode-indicator .admin-dot {
            width: 8px; height: 8px; border-radius: 50%; background: #4ec06a;
            box-shadow: 0 0 6px #4ec06a;
        }
        @keyframes adminPulse {
            0%, 100% { box-shadow: 0 2px 6px rgba(26, 58, 107, 0.3); }
            50% { box-shadow: 0 2px 14px rgba(26, 58, 107, 0.5); }
        }
        body.dark-theme .admin-mode-indicator {
            background: linear-gradient(135deg, #4a82c7, #3a6ba8);
        }

        .btn-logout {
            background: transparent; color: var(--danger); border: 1.5px solid var(--danger);
        }
        .btn-logout:hover { background: var(--danger); color: white; }

        .filter-bar {
            display: flex; flex-wrap: wrap; gap: 10px; align-items: center;
            padding: 10px 14px; background: var(--accent-lighter);
            border: 1px solid var(--border); border-radius: 12px;
            transition: background 0.3s ease, border-color 0.3s ease; flex-shrink: 0;
        }
        .filter-bar .filter-label { font-weight: 600; font-size: 0.75rem; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.4px; }

        .search-wrapper { position: relative; flex: 1 1 220px; min-width: 180px; }
        .search-wrapper input {
            width: 100%; padding: 8px 12px 8px 34px; border-radius: 30px;
            border: 1.5px solid var(--border); background: var(--table-bg);
            color: var(--text-primary); font-size: 0.82rem; font-family: inherit; outline: none;
            transition: border-color 0.15s, box-shadow 0.15s, background 0.3s;
        }
        .search-wrapper input:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(74, 130, 199, 0.18); }
        .search-wrapper .search-icon { position: absolute; left: 12px; top: 50%; transform: translateY(-50%); color: var(--text-muted); font-size: 0.9rem; pointer-events: none; }
        .search-wrapper .clear-search {
            position: absolute; right: 8px; top: 50%; transform: translateY(-50%);
            background: none; border: none; color: var(--text-muted); cursor: pointer;
            font-size: 1rem; padding: 3px 6px; border-radius: 50%; display: none; font-family: inherit;
        }
        .search-wrapper .clear-search.visible { display: block; }
        .search-wrapper .clear-search:hover { color: var(--danger); background: var(--accent-light); }

        .teacher-select {
            flex: 0 1 220px; min-width: 160px; padding: 8px 12px; border-radius: 30px;
            border: 1.5px solid var(--border); background: var(--table-bg);
            color: var(--text-primary); font-size: 0.82rem; font-family: inherit;
            outline: none; cursor: pointer; transition: border-color 0.15s, background 0.3s;
        }
        .teacher-select:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(74, 130, 199, 0.18); }
        .teacher-select option { background: var(--table-bg); color: var(--text-primary); }

        .filter-count { font-size: 0.76rem; color: var(--text-muted); font-weight: 600; white-space: nowrap; }

        .substitute-toggle {
            display: inline-flex; align-items: center; gap: 6px; padding: 6px 12px;
            border-radius: 30px; border: 1.5px solid var(--substitute-border);
            background: var(--substitute-bg); color: var(--substitute-text);
            font-weight: 600; font-size: 0.76rem; cursor: pointer;
            user-select: none; transition: all 0.15s; white-space: nowrap;
        }
        .substitute-toggle:hover { filter: brightness(1.05); }
        .substitute-toggle input { cursor: pointer; accent-color: var(--substitute-border); width: 14px; height: 14px; }

        .day-filter { display: flex; gap: 6px; flex-wrap: wrap; flex-shrink: 0; }
        .day-btn {
            padding: 6px 14px; border-radius: 40px; border: 1.5px solid var(--border);
            background: var(--table-bg); color: var(--text-primary);
            font-weight: 600; font-size: 0.76rem; cursor: pointer;
            transition: all 0.15s; font-family: inherit;
        }
        .day-btn.active { background: var(--accent); color: white; border-color: var(--accent); }
        .day-btn:hover:not(.active) { background: var(--accent-light); }

        .table-wrapper {
            flex: 1 1 auto; min-height: 0; overflow: auto;
            border-radius: 12px; border: 1px solid var(--border);
            background: var(--table-bg); box-shadow: var(--shadow-table);
            transition: background 0.3s, border-color 0.3s;
        }

        table { width: 100%; border-collapse: collapse; font-size: 0.78rem; min-width: 1300px; }

        thead th {
            background: var(--th-bg); color: white; font-weight: 600;
            padding: 8px 5px; text-align: center; font-size: 0.74rem;
            border-right: 1px solid rgba(255,255,255,0.15);
            position: sticky; z-index: 5;
        }
        thead tr:first-child th { top: 0; }
        thead tr:last-child th  { top: 34px; font-size: 0.68rem; padding: 4px 2px; }
        thead th:last-child { border-right: none; }

        td {
            padding: 6px 5px; border: 1px solid var(--border-light); vertical-align: top;
            background: var(--table-bg); transition: background 0.1s;
            min-width: 78px; word-break: break-word; line-height: 1.25; color: var(--text-primary);
        }
        tr:nth-child(even) td { background: var(--table-row-even); }
        tr:hover td { background: var(--table-row-hover); }

        .teacher-name {
            font-weight: 700; color: var(--text-primary); background: var(--teacher-bg) !important;
            position: sticky; left: 0; z-index: 4; min-width: 140px;
            border-right: 2px solid var(--accent); font-size: 0.74rem;
        }
        .room-cell {
            background: var(--room-bg) !important; font-weight: 500; color: var(--text-secondary);
            font-size: 0.7rem; text-align: center; min-width: 48px;
            position: sticky; left: 140px; z-index: 3;
        }
        .class-cell {
            background: var(--room-bg) !important; font-weight: 500; color: var(--text-secondary);
            font-size: 0.7rem; text-align: center; min-width: 44px;
            position: sticky; left: 188px; z-index: 3;
        }

        .lesson-cell {
            position: relative; transition: background 0.15s;
            font-size: 0.74rem; color: var(--text-primary); font-weight: 500;
        }
        .lesson-cell.has-substitute { background: var(--substitute-bg) !important; border-color: var(--substitute-border) !important; }
        tr:nth-child(even) .lesson-cell.has-substitute { background: var(--substitute-bg-even) !important; }
        tr:hover .lesson-cell.has-substitute { background: var(--substitute-bg-hover) !important; }
        .lesson-cell.has-substitute .lesson-text { color: var(--substitute-text); font-weight: 700; }
        .lesson-cell .substitute-note {
            display: block; font-size: 0.62rem; font-weight: 600; font-style: italic;
            color: var(--substitute-text); margin-top: 1px; opacity: 0.9;
        }
        .lesson-cell .substitute-icon { position: absolute; top: 1px; right: 2px; font-size: 0.62rem; opacity: 0.8; }

        body.edit-mode .lesson-cell { cursor: pointer; }
        body.edit-mode .lesson-cell:hover {
            background: var(--edit-hover) !important;
            outline: 2px solid var(--edit-outline); outline-offset: -2px;
        }
        body.edit-mode .lesson-cell.has-substitute:hover {
            background: var(--substitute-bg-hover) !important;
            outline: 2px solid var(--substitute-border);
        }

        .lesson-cell.recently-changed { animation: highlightChange 1.5s ease; }
        @keyframes highlightChange {
            0%   { background: #d4f4dd; }
            100% { background: inherit; }
        }

        body:not(.edit-mode) .lesson-cell { cursor: default; }
        body:not(.edit-mode) .lesson-cell:hover { background: inherit !important; outline: none !important; }
        body:not(.edit-mode) .lesson-cell.has-substitute:hover { background: var(--substitute-bg) !important; }

        .lesson-cell.editing {
            padding: 0; outline: 2px solid var(--accent); outline-offset: -2px;
            background: var(--table-bg) !important; cursor: text;
            position: relative; z-index: 10;
        }

        .cell-editor { padding: 5px; display: flex; flex-direction: column; gap: 3px; min-width: 200px; }
        .cell-editor input[type="text"] {
            width: 100%; padding: 4px 6px; font-size: 0.74rem; font-family: inherit;
            border: 1px solid var(--border); border-radius: 5px;
            background: var(--table-bg); color: var(--text-primary); outline: none;
        }
        .cell-editor input[type="text"]:focus { border-color: var(--accent); box-shadow: 0 0 0 2px rgba(74, 130, 199, 0.2); }
        .cell-editor input[type="text"].substitute-input {
            border-color: var(--substitute-border); background: var(--substitute-bg);
            color: var(--substitute-text); font-weight: 600;
        }
        .cell-editor input[type="text"].substitute-input:focus { border-color: var(--substitute-border); box-shadow: 0 0 0 2px rgba(212, 160, 23, 0.25); }

        .cell-editor .substitute-label {
            display: flex; align-items: center; justify-content: space-between;
            font-size: 0.6rem; color: var(--substitute-text); font-weight: 700;
            text-transform: uppercase; letter-spacing: 0.2px; margin-top: 1px;
        }
        .cell-editor .free-counter { font-size: 0.6rem; color: var(--free-color); font-weight: 700; text-transform: none; letter-spacing: 0; }
        .cell-editor .free-counter.all-busy { color: var(--busy-color); }

        .cell-editor .substitute-row { display: flex; align-items: center; gap: 4px; position: relative; }
        .cell-editor .substitute-row .substitute-input { flex: 1; }
        .cell-editor .substitute-row .clear-substitute {
            background: none; border: 1px solid var(--border); color: var(--text-muted);
            border-radius: 4px; padding: 3px 5px; font-size: 0.66rem;
            cursor: pointer; font-family: inherit; transition: all 0.15s;
        }
        .cell-editor .substitute-row .clear-substitute:hover { background: var(--danger); color: white; border-color: var(--danger); }

        .free-teachers-dropdown {
            position: absolute; top: calc(100% + 3px); left: 0; right: 0;
            background: var(--modal-bg); border: 1.5px solid var(--substitute-border);
            border-radius: 6px; box-shadow: 0 8px 24px rgba(0,0,0,0.2);
            max-height: 200px; overflow-y: auto; z-index: 100; font-size: 0.74rem;
        }
        .free-teachers-dropdown.hidden { display: none; }
        .free-teachers-dropdown .dropdown-header {
            padding: 5px 8px; font-size: 0.64rem; font-weight: 700;
            text-transform: uppercase; letter-spacing: 0.4px; color: var(--text-muted);
            background: var(--accent-lighter); border-bottom: 1px solid var(--border);
            position: sticky; top: 0; z-index: 1;
            display: flex; justify-content: space-between; align-items: center;
        }
        .free-teachers-dropdown .dropdown-header .refresh-btn {
            background: none; border: none; cursor: pointer; font-size: 0.85rem;
            color: var(--accent); padding: 1px 3px; border-radius: 4px; font-family: inherit;
        }
        .free-teachers-dropdown .dropdown-header .refresh-btn:hover { background: var(--accent-light); }
        .free-teachers-dropdown .dropdown-item {
            padding: 5px 8px; cursor: pointer; display: flex; align-items: center; gap: 6px;
            color: var(--text-primary); border-bottom: 1px solid var(--border-light);
            transition: background 0.1s;
        }
        .free-teachers-dropdown .dropdown-item:last-child { border-bottom: none; }
        .free-teachers-dropdown .dropdown-item:hover,
        .free-teachers-dropdown .dropdown-item.highlighted { background: var(--substitute-bg); color: var(--substitute-text); }
        .free-teachers-dropdown .dropdown-item .status-dot {
            width: 7px; height: 7px; border-radius: 50%; background: var(--free-color); flex-shrink: 0;
        }
        .free-teachers-dropdown .dropdown-item .teacher-info { flex: 1; display: flex; flex-direction: column; }
        .free-teachers-dropdown .dropdown-item .teacher-name-item { font-weight: 600; }
        .free-teachers-dropdown .dropdown-item .teacher-meta { font-size: 0.62rem; color: var(--text-muted); font-style: italic; }
        .free-teachers-dropdown .empty-item {
            padding: 10px 8px; text-align: center; color: var(--busy-color);
            font-style: italic; font-size: 0.7rem; font-weight: 600;
        }

        .cell-editor .editor-hint { font-size: 0.6rem; color: var(--text-muted); font-style: italic; }
        .empty-lesson { color: var(--empty-color); font-style: italic; font-weight: 400; }

        .footer-note {
            text-align: center; font-size: 0.7rem; color: var(--text-muted);
            border-top: 1px solid var(--border); padding-top: 8px; flex-shrink: 0;
        }
        .status-badge {
            font-size: 0.66rem; background: var(--success); color: white;
            padding: 2px 8px; border-radius: 20px; font-weight: 500; transition: background 0.3s;
        }

        .modal-overlay {
            position: fixed; inset: 0; background: rgba(8, 20, 40, 0.55);
            backdrop-filter: blur(4px); z-index: 3000; display: none;
            align-items: center; justify-content: center; padding: 16px;
        }
        .modal-overlay.visible { display: flex; }
        .modal {
            background: var(--modal-bg); border-radius: 16px; padding: 22px 24px;
            width: 100%; max-width: 400px; max-height: calc(100vh - 32px);
            overflow-y: auto; box-shadow: 0 24px 48px rgba(0,0,0,0.3);
            animation: modalIn 0.25s ease; color: var(--text-primary);
        }
        .modal.modal-wide {
            max-width: 900px; padding: 20px 22px; max-height: calc(100vh - 32px);
            display: flex; flex-direction: column;
        }
        .modal.modal-narrow { max-width: 460px; }

        @keyframes modalIn {
            from { transform: translateY(20px); opacity: 0; }
            to   { transform: translateY(0);    opacity: 1; }
        }

        .modal h2 { color: var(--text-primary); font-size: 1.1rem; margin-bottom: 5px; }
        .modal p  { color: var(--text-secondary); font-size: 0.82rem; margin-bottom: 14px; }
        .modal label { display: block; font-weight: 600; color: var(--accent); font-size: 0.78rem; margin-bottom: 5px; }

        .modal input[type="password"], .modal input[type="text"], .modal input[type="time"], .modal input[type="date"], .modal select {
            width: 100%; padding: 9px 12px; border-radius: 10px;
            border: 1.5px solid var(--border); background: var(--table-bg);
            color: var(--text-primary); font-size: 0.85rem; outline: none;
            transition: border 0.15s; margin-bottom: 6px; font-family: inherit;
        }
        .modal input:focus, .modal select:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(74, 130, 199, 0.18); }
        .modal .error-msg { color: var(--danger); font-size: 0.76rem; font-weight: 600; min-height: 16px; margin-bottom: 8px; display: block; }
        .modal-actions { display: flex; gap: 8px; justify-content: flex-end; margin-top: 4px; }
        .modal-actions .btn { padding: 8px 18px; }

        #consoleModal .modal.modal-wide {
            max-width: min(1400px, 98vw); width: 98vw;
            height: calc(100vh - 20px); max-height: calc(100vh - 20px);
            padding: 14px 18px; display: flex; flex-direction: column; border-radius: 14px;
        }
        #consoleModal { z-index: 2000; }
        #teacherEditModal { z-index: 3100; }
        #myScheduleModal { z-index: 2500; }
        #bellsModal { z-index: 3100; }
        #dictManageModal { z-index: 3200; }

        .console-header {
            display: flex; align-items: center; justify-content: space-between;
            padding-bottom: 8px; border-bottom: 1.5px solid var(--border);
            margin-bottom: 8px; flex-shrink: 0;
        }
        .console-header h2 { margin: 0; display: flex; align-items: center; gap: 6px; font-size: 1rem; }

        .console-tabs {
            display: flex; gap: 2px; border-bottom: 1.5px solid var(--border);
            margin-bottom: 10px; flex-shrink: 0; overflow-x: auto;
        }
        .console-tab {
            padding: 6px 12px; border: none; background: none;
            color: var(--text-secondary); font-weight: 600; font-size: 0.72rem;
            cursor: pointer; border-bottom: 2.5px solid transparent;
            margin-bottom: -1.5px; white-space: nowrap; font-family: inherit;
            transition: all 0.15s;
        }
        .console-tab.active { color: var(--accent); border-bottom-color: var(--accent); }
        .console-tab:hover:not(.active) { color: var(--accent); }

        .console-tab.locked { position: relative; padding-right: 26px; }
        .console-tab.locked::after {
            content: "🔒"; position: absolute; right: 6px; top: 50%;
            transform: translateY(-50%); font-size: 0.7rem; opacity: 0.7;
        }

        .console-body {
            flex: 1 1 auto; overflow-y: auto; overflow-x: hidden;
            padding-right: 4px; min-height: 0;
        }
        .console-tab-content { display: none; }
        .console-tab-content.active { display: block; }

        .console-toolbar {
            display: flex; gap: 6px; flex-wrap: wrap;
            margin-bottom: 8px; align-items: center;
        }
        .console-toolbar .spacer { flex: 1; }

        .admin-table { width: 100%; border-collapse: collapse; font-size: 0.72rem; }
        .admin-table th {
            background: var(--accent-light); color: var(--text-primary);
            padding: 5px 6px; text-align: left; font-size: 0.66rem; font-weight: 700;
            border-bottom: 1.5px solid var(--border); position: sticky;
            top: 0; z-index: 2; text-transform: uppercase; letter-spacing: 0.3px;
        }
        .admin-table td {
            padding: 3px 6px; border-bottom: 1px solid var(--border-light);
            background: transparent; color: var(--text-primary);
            vertical-align: middle; font-size: 0.72rem;
        }
        .admin-table tr:hover td { background: var(--table-row-hover); }
        .admin-table input {
            width: 100%; padding: 3px 6px; border-radius: 5px;
            border: 1px solid var(--border); background: var(--table-bg);
            color: var(--text-primary); font-size: 0.72rem; font-family: inherit;
            outline: none; box-sizing: border-box;
        }
        .admin-table input:focus { border-color: var(--accent); box-shadow: 0 0 0 2px rgba(74, 130, 199, 0.15); }

        .admin-actions { display: flex; gap: 2px; white-space: nowrap; }

        .icon-btn {
            background: none; border: none; cursor: pointer; font-size: 0.82rem;
            padding: 2px 4px; border-radius: 4px; color: var(--text-secondary);
            transition: all 0.15s; font-family: inherit; line-height: 1;
        }
        .icon-btn:hover { background: var(--accent-light); color: var(--accent); }
        .icon-btn.danger:hover { background: #ffe5ea; color: var(--danger); }

        .history-list { display: flex; flex-direction: column; gap: 4px; }
        .history-item {
            display: flex; gap: 8px; padding: 6px 10px; border-radius: 8px;
            background: var(--accent-lighter); border: 1px solid var(--border);
            font-size: 0.72rem; align-items: center; transition: all 0.15s; line-height: 1.35;
        }
        .history-item:hover { border-color: var(--accent); }
        .history-item.undone { opacity: 0.5; text-decoration: line-through; background: var(--table-bg); }
        .history-item.substitute-entry { background: var(--substitute-bg); border-color: var(--substitute-border); }
        .history-time { color: var(--text-muted); font-size: 0.64rem; font-weight: 600; white-space: nowrap; min-width: 84px; }
        .history-content { flex: 1; min-width: 0; }
        .history-content .who { font-weight: 700; color: var(--accent); }
        .history-content .detail { color: var(--text-secondary); margin-top: 1px; font-size: 0.68rem; }
        .history-content .old-val { color: var(--danger); text-decoration: line-through; font-style: italic; }
        .history-content .new-val { color: var(--success); font-weight: 600; }
        .history-content .subst-val { color: var(--substitute-text); font-weight: 700; }
        .history-empty { text-align: center; padding: 20px 16px; color: var(--text-muted); font-style: italic; font-size: 0.8rem; }

        .security-block {
            padding: 10px 12px; background: var(--accent-lighter);
            border-radius: 10px; border: 1px solid var(--border); margin-bottom: 8px;
        }
        .security-block h3 {
            font-size: 0.82rem; color: var(--text-primary); margin-bottom: 6px;
            display: flex; align-items: center; gap: 6px;
        }
        .security-block p { color: var(--text-secondary); font-size: 0.7rem; margin-bottom: 8px; line-height: 1.4; }

        #security-tab-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; align-items: start; }
        @media (max-width: 800px) { #security-tab-grid { grid-template-columns: 1fr; } }

        .password-field { position: relative; margin-bottom: 6px; }
        .password-field label {
            display: block; font-size: 0.66rem; font-weight: 600;
            color: var(--text-secondary); margin-bottom: 3px;
            text-transform: uppercase; letter-spacing: 0.3px;
        }
        .password-field input {
            width: 100%; padding: 6px 32px 6px 10px; border-radius: 6px;
            border: 1.5px solid var(--border); background: var(--table-bg);
            color: var(--text-primary); font-size: 0.78rem; font-family: inherit;
            outline: none; transition: border-color 0.15s; box-sizing: border-box;
        }
        .password-field input:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(74, 130, 199, 0.15); }
        .password-field .toggle-visibility {
            position: absolute; right: 4px; bottom: 4px; background: none;
            border: none; color: var(--text-muted); cursor: pointer;
            font-size: 0.85rem; padding: 2px 5px; border-radius: 4px;
            font-family: inherit; line-height: 1;
        }
        .password-field .toggle-visibility:hover { background: var(--accent-light); color: var(--accent); }

        .password-strength { height: 3px; border-radius: 2px; background: var(--border); margin-top: 4px; overflow: hidden; }
        .password-strength .bar { height: 100%; width: 0%; transition: width 0.3s ease, background 0.3s ease; background: var(--danger); }
        .password-strength.weak .bar { width: 25%; background: #d05050; }
        .password-strength.medium .bar { width: 55%; background: var(--warning); }
        .password-strength.strong .bar { width: 100%; background: var(--success); }

        .password-strength-label { font-size: 0.62rem; margin-top: 2px; color: var(--text-muted); font-weight: 600; }
        .password-strength-label.weak { color: #d05050; }
        .password-strength-label.medium { color: var(--warning); }
        .password-strength-label.strong { color: var(--success); }

        .password-msg { font-size: 0.7rem; font-weight: 600; margin-top: 6px; min-height: 14px; }
        .password-msg.error { color: var(--danger); }
        .password-msg.success { color: var(--success); }

        .password-info {
            padding: 6px 10px; background: var(--substitute-bg);
            border: 1px solid var(--substitute-border); border-radius: 6px;
            font-size: 0.68rem; color: var(--substitute-text);
            margin-top: 6px; line-height: 1.4;
        }

        #io-tab-grid { display: grid; grid-template-columns: 1.2fr 1fr; gap: 10px; align-items: start; }
        @media (max-width: 900px) { #io-tab-grid { grid-template-columns: 1fr; } }

        .io-block {
            padding: 10px 12px; background: var(--accent-lighter);
            border-radius: 10px; border: 1px solid var(--border); margin-bottom: 8px;
        }
        .io-block.io-excel { border: 1.5px solid var(--accent); }
        .io-block strong { color: var(--text-primary); font-size: 0.8rem; display: block; }
        .io-block p { color: var(--text-secondary); font-size: 0.68rem; margin: 3px 0 6px; line-height: 1.35; }

        .import-preview {
            margin-top: 6px; border: 1px solid var(--border); border-radius: 8px;
            overflow: hidden; max-height: 180px; overflow-y: auto;
        }
        .import-preview table { width: 100%; min-width: auto; font-size: 0.66rem; border-collapse: collapse; }
        .import-preview th {
            background: var(--accent-light); color: var(--text-primary);
            padding: 4px 6px; font-size: 0.62rem; font-weight: 700;
            text-align: left; position: sticky; top: 0; z-index: 1;
            border-bottom: 1px solid var(--border);
        }
        .import-preview td { padding: 3px 6px; border-bottom: 1px solid var(--border-light); color: var(--text-primary); font-size: 0.66rem; }
        .import-preview .empty-val { color: var(--empty-color); font-style: italic; }

        #consoleModal .btn { padding: 5px 10px; font-size: 0.7rem; gap: 4px; }
        #consoleModal .btn-sm { padding: 4px 9px; font-size: 0.68rem; }
        #consoleModal .teacher-select { padding: 6px 10px; font-size: 0.72rem; min-width: 140px; }

        #bulkStats { font-size: 0.7rem !important; line-height: 1.5; }
        #bulkStats div { margin-bottom: 2px; }

        #storageInfo { font-size: 0.72rem !important; line-height: 1.6; display: grid; grid-template-columns: 1fr 1fr; gap: 4px 12px; }

        .my-sched-controls {
            display: flex; flex-wrap: wrap; gap: 12px; align-items: flex-end;
            padding: 12px 14px; background: var(--accent-lighter);
            border-radius: 12px; border: 1px solid var(--border);
            margin-bottom: 14px; flex-shrink: 0;
        }
        .my-sched-control { display: flex; flex-direction: column; gap: 4px; flex: 1 1 200px; min-width: 160px; }
        .my-sched-control label {
            font-size: 0.72rem; font-weight: 600; color: var(--text-secondary);
            text-transform: uppercase; letter-spacing: 0.3px;
        }
        .my-sched-control .teacher-select { flex: none; width: 100%; margin: 0; }
        .my-sched-control > div { display: flex; gap: 6px; }

        .my-sched-day { margin-bottom: 16px; border-radius: 12px; border: 1px solid var(--border); background: var(--table-bg); overflow: hidden; }
        .my-sched-day-header {
            padding: 10px 14px; background: var(--th-bg); color: white;
            font-weight: 700; font-size: 0.9rem;
            display: flex; justify-content: space-between; align-items: center;
        }
        .my-sched-day-header .substitutions-count {
            font-size: 0.75rem; background: var(--substitute-badge-bg);
            color: var(--substitute-badge-text); padding: 3px 8px;
            border-radius: 12px; font-weight: 600;
        }
        .my-sched-lessons { display: flex; flex-direction: column; }
        .my-sched-lesson {
            display: grid; grid-template-columns: 44px 100px 1fr 80px 80px;
            gap: 10px; padding: 10px 14px;
            border-bottom: 1px solid var(--border-light);
            align-items: center; font-size: 0.82rem; transition: background 0.15s;
        }
        .my-sched-lesson:last-child { border-bottom: none; }
        .my-sched-lesson:hover { background: var(--table-row-hover); }
        .my-sched-lesson .num { font-weight: 700; color: var(--accent); font-size: 1rem; text-align: center; }
        .my-sched-lesson .time { font-size: 0.72rem; color: var(--text-muted); font-weight: 600; text-align: center; }
        .my-sched-lesson .subject { font-weight: 600; color: var(--text-primary); font-size: 0.88rem; }
        .my-sched-lesson .class-name { font-weight: 600; color: var(--accent); font-size: 0.8rem; text-align: center; }
        .my-sched-lesson .room { font-size: 0.75rem; color: var(--text-secondary); text-align: center; font-weight: 600; }
        .my-sched-lesson.empty { background: var(--table-row-even); color: var(--text-muted); font-style: italic; }
        .my-sched-lesson.empty .subject { color: var(--text-muted); font-weight: 400; font-style: italic; }
        .my-sched-lesson.has-substitute { background: var(--substitute-bg); border-left: 4px solid var(--substitute-border); }
        .my-sched-lesson.has-substitute:hover { background: var(--substitute-bg-hover); }
        .my-sched-lesson.has-substitute .subject { color: var(--substitute-text); }
        .my-sched-lesson .subst-note {
            display: block; font-size: 0.72rem; font-weight: 600;
            font-style: italic; color: var(--substitute-text); margin-top: 2px;
        }
        .my-sched-lesson .subst-icon { margin-right: 4px; }

        .my-sched-replacements {
            margin-top: 16px; padding: 14px;
            background: var(--substitute-bg);
            border: 1.5px solid var(--substitute-border); border-radius: 12px;
        }
        .my-sched-replacements h3 {
            font-size: 0.95rem; color: var(--substitute-text);
            margin-bottom: 10px; display: flex; align-items: center; gap: 8px;
        }
        .my-sched-replacement-row {
            display: grid; grid-template-columns: 44px 100px 1fr 1fr;
            gap: 10px; padding: 8px 0;
            border-bottom: 1px dashed var(--substitute-border);
            font-size: 0.82rem; align-items: center;
        }
        .my-sched-replacement-row:last-child { border-bottom: none; }
        .my-sched-replacement-row .num { font-weight: 700; color: var(--substitute-text); text-align: center; }
        .my-sched-replacement-row .time { font-size: 0.72rem; color: var(--text-muted); font-weight: 600; text-align: center; }
        .my-sched-replacement-row .who { font-weight: 600; color: var(--substitute-text); }
        .my-sched-empty { text-align: center; padding: 40px 20px; color: var(--text-muted); font-size: 0.9rem; font-style: italic; }
        .my-sched-hint { text-align: center; padding: 20px; color: var(--text-secondary); font-size: 0.85rem; }

        @media print {
            body * { visibility: hidden; }
            #myScheduleModal, #myScheduleModal * { visibility: visible; }
            #myScheduleModal {
                position: absolute; inset: 0; background: white;
                display: block !important; padding: 20px;
            }
            .modal-wide {
                max-width: 100% !important; max-height: none !important;
                box-shadow: none !important; padding: 0 !important;
            }
            .console-header .btn, .my-sched-controls, #closeMyScheduleBtn { display: none !important; }
            .my-sched-day { page-break-inside: avoid; }
            .my-sched-lesson { padding: 6px 10px; }
            .my-sched-lesson .num { color: #000; }
            .my-sched-day-header {
                background: #1a3a6b !important; color: white !important;
                -webkit-print-color-adjust: exact; print-color-adjust: exact;
            }
            .my-sched-lesson.has-substitute {
                background: #fff3cd !important;
                -webkit-print-color-adjust: exact; print-color-adjust: exact;
            }
        }

        @media (max-width: 700px) {
            .my-sched-lesson {
                grid-template-columns: 36px 1fr;
                grid-template-areas: "num subject" "num class-room";
                gap: 4px 10px;
            }
            .my-sched-lesson .num { grid-area: num; align-self: center; }
            .my-sched-lesson .subject { grid-area: subject; }
            .my-sched-lesson .time, .my-sched-lesson .class-name, .my-sched-lesson .room { font-size: 0.75rem; }
            .my-sched-lesson .time { display: none; }
            .my-sched-lesson .class-name, .my-sched-lesson .room { grid-area: class-room; display: inline; }
            .my-sched-lesson .class-name::after { content: " · "; }
        }

        .no-data-row td {
            text-align: center; padding: 30px 20px; color: var(--text-muted);
            font-size: 0.9rem; font-style: italic; background: var(--table-bg) !important;
        }

        .console-btn-wrapper { position: relative; display: inline-flex; }
        .console-badge {
            position: absolute; top: -4px; right: -4px; background: var(--danger);
            color: white; font-size: 0.6rem; font-weight: 700;
            min-width: 16px; height: 16px; border-radius: 8px;
            display: flex; align-items: center; justify-content: center;
            padding: 0 4px; border: 2px solid var(--bg-container); pointer-events: none;
        }
        .console-badge.hidden { display: none; }
        .console-badge.subst { background: var(--substitute-badge-bg); color: var(--substitute-badge-text); }

        .theme-toggle {
            position: relative; width: 56px; height: 30px; border-radius: 20px;
            background: var(--accent-light); border: 1.5px solid var(--border);
            cursor: pointer; display: inline-flex; align-items: center; padding: 2px;
            transition: background 0.3s, border-color 0.3s; flex-shrink: 0;
        }
        .theme-toggle .knob {
            width: 22px; height: 22px; border-radius: 50%; background: var(--accent);
            display: flex; align-items: center; justify-content: center;
            font-size: 0.7rem; transition: transform 0.3s cubic-bezier(.4,0,.2,1), background 0.3s;
            transform: translateX(0);
        }
        body.dark-theme .theme-toggle .knob { transform: translateX(26px); background: #f0c040; }
        .theme-toggle .knob .icon-sun  { display: block; }
        .theme-toggle .knob .icon-moon { display: none; }
        body.dark-theme .theme-toggle .knob .icon-sun  { display: none; }
        body.dark-theme .theme-toggle .knob .icon-moon { display: block; }

        .history-panel {
            border-radius: 12px; border: 1px solid var(--border);
            background: var(--table-bg); padding: 12px 14px; display: none;
            flex-shrink: 0; max-height: 260px; overflow: hidden;
        }
        .history-panel.visible { display: flex; flex-direction: column; }
        .history-panel-header {
            display: flex; align-items: center; justify-content: space-between;
            margin-bottom: 10px; flex-wrap: wrap; gap: 8px; flex-shrink: 0;
        }
        .history-panel-header h3 { font-size: 0.9rem; color: var(--text-primary); display: flex; align-items: center; gap: 6px; }
        .history-panel-list { overflow-y: auto; display: flex; flex-direction: column; gap: 5px; flex: 1; }

        .toast {
            position: fixed; bottom: 20px; right: 20px; background: var(--success);
            color: white; padding: 10px 16px; border-radius: 10px;
            box-shadow: 0 8px 24px rgba(0,0,0,0.25); font-weight: 600;
            font-size: 0.8rem; z-index: 5000; animation: toastIn 0.3s ease; max-width: 320px;
        }
        .toast.error { background: var(--danger); }
        .toast.warning { background: var(--warning); color: #1a1a1a; }
        .toast.substitute { background: var(--substitute-badge-bg); color: var(--substitute-badge-text); }
        @keyframes toastIn {
            from { transform: translateY(20px); opacity: 0; }
            to   { transform: translateY(0);    opacity: 1; }
        }

        @media (max-width: 900px) {
            .app-container { padding: 12px; }
            .school-title h1 { font-size: 1.05rem; }
            .school-title p, .school-title .subtitle { font-size: 0.7rem; }
            .teacher-name { min-width: 110px; font-size: 0.68rem; }
            .room-cell { left: 110px; }
            .class-cell { left: 155px; }
            .filter-bar { padding: 8px 10px; }
        }

        /* ==================== ВКЛАДКА «ЖУРНАЛ ЗАМЕН» ==================== */
        #substitutionsPanel {
            display: none;
            flex-direction: column;
            gap: 12px;
            flex: 1 1 auto;
            min-height: 0;
            overflow: hidden;
        }
        body.mode-substitutions #substitutionsPanel { display: flex; }
        body.mode-substitutions #tableWrapper,
        body.mode-substitutions #dayFilter,
        body.mode-substitutions .filter-bar,
        body.mode-substitutions #scheduleHint,
        body.mode-substitutions #historyPanel { display: none !important; }

        .subst-toolbar {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            align-items: center;
            padding: 10px 14px;
            background: var(--accent-lighter);
            border: 1px solid var(--border);
            border-radius: 12px;
            flex-shrink: 0;
        }
        .subst-toolbar .filter-label {
            font-weight: 600; font-size: 0.75rem; color: var(--text-secondary);
            text-transform: uppercase; letter-spacing: 0.4px;
        }
        .subst-toolbar select,
        .subst-toolbar input[type="text"],
        .subst-toolbar input[type="month"],
        .subst-toolbar input[type="date"] {
            padding: 7px 12px; border-radius: 30px;
            border: 1.5px solid var(--border); background: var(--table-bg);
            color: var(--text-primary); font-size: 0.8rem; font-family: inherit;
            outline: none; transition: border-color 0.15s;
        }
        .subst-toolbar select:focus,
        .subst-toolbar input:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(74,130,199,0.18); }

        .subst-stat-card {
            padding: 10px 14px;
            background: var(--table-bg);
            border: 1px solid var(--border);
            border-radius: 10px;
            box-shadow: var(--shadow-table);
        }
        .subst-stat-card .label {
            font-size: 0.68rem; color: var(--text-muted); font-weight: 600;
            text-transform: uppercase; letter-spacing: 0.3px;
        }
        .subst-stat-card .value {
            font-size: 1.4rem; font-weight: 700; color: var(--accent);
            margin-top: 3px; line-height: 1.1;
        }
        .subst-stat-card .value.subst { color: var(--substitute-text); }

        .subst-body {
            display: flex;
            flex-direction: column;
            gap: 12px;
            min-height: 0;
            overflow-y: auto;
            overflow-x: hidden;
            padding-right: 4px;
            position: relative;
        }

        .subst-section {
            background: var(--table-bg);
            border: 1px solid var(--border);
            border-radius: 12px;
            overflow: hidden;
            box-shadow: var(--shadow-table);
            display: flex;
            flex-direction: column;
            min-height: 0;
            flex: 1 1 auto;
        }
        .subst-section > div {
            overflow: visible;
            min-height: 0;
            position: relative;
        }

        .subst-section-header {
            padding: 10px 14px;
            background: var(--th-bg);
            color: white;
            font-weight: 700;
            font-size: 0.85rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 8px;
            flex-wrap: wrap;
            position: sticky;
            top: 0;
            z-index: 5;
        }
        .subst-section-header .count-badge {
            font-size: 0.72rem;
            background: rgba(255,255,255,0.2);
            padding: 3px 10px;
            border-radius: 20px;
            font-weight: 600;
        }

        .subst-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.78rem;
        }
        .subst-table th {
            background: var(--accent-light);
            color: var(--text-primary);
            padding: 8px 10px;
            text-align: left;
            font-size: 0.7rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            border-bottom: 1.5px solid var(--border);
            position: sticky;
            top: 0;
            z-index: 3;
            box-shadow: 0 1px 0 var(--border);
            background-clip: padding-box;
        }
        .subst-table td {
            padding: 8px 10px;
            border-bottom: 1px solid var(--border-light);
            color: var(--text-primary);
            vertical-align: middle;
        }
        .subst-table tr:hover td { background: var(--table-row-hover); }
        .subst-table td.actions { white-space: nowrap; text-align: center; width: 80px; }
        .subst-table .empty-row td {
            text-align: center; padding: 30px 20px; color: var(--text-muted);
            font-style: italic; font-size: 0.85rem;
        }

        .subst-type-badge {
            display: inline-block;
            font-size: 0.66rem;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 12px;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            white-space: nowrap;
        }
        .subst-type-badge.standard {
            background: var(--accent-light);
            color: var(--accent);
            border: 1px solid var(--border);
        }
        .subst-type-badge.extra {
            background: var(--substitute-badge-bg);
            color: var(--substitute-badge-text);
            border: 1px solid var(--substitute-border);
        }

        .subst-history-panel {
            border-radius: 12px;
            border: 1px solid var(--border);
            background: var(--table-bg);
            padding: 12px 14px;
            display: none;
            flex-shrink: 0;
            max-height: 260px;
            overflow: hidden;
            margin-top: 4px;
        }
        .subst-history-panel.visible { display: flex; flex-direction: column; }
        .subst-history-panel .header {
            display: flex; align-items: center; justify-content: space-between;
            margin-bottom: 10px; flex-wrap: wrap; gap: 8px; flex-shrink: 0;
        }
        .subst-history-panel .header h3 {
            font-size: 0.9rem; color: var(--text-primary);
            display: flex; align-items: center; gap: 6px;
        }
        .subst-history-list {
            overflow-y: auto; display: flex; flex-direction: column;
            gap: 5px; flex: 1;
        }

        .subst-add-panel {
            padding: 8px 12px;
            background: var(--accent-lighter);
            border: 1px solid var(--border);
            border-radius: 12px;
            flex-shrink: 0;
        }
        .subst-add-title {
            font-weight: 700;
            color: var(--text-primary);
            font-size: 0.85rem;
            margin-bottom: 6px;
        }
        .subst-add-grid {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 6px 8px;
            align-items: end;
        }
        .subst-add-field { display: flex; flex-direction: column; gap: 2px; }
        .subst-add-field label {
            font-size: 0.66rem;
            font-weight: 600;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.3px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .subst-add-field input,
        .subst-add-field select {
            padding: 5px 8px;
            border-radius: 6px;
            border: 1.5px solid var(--border);
            background: var(--table-bg);
            color: var(--text-primary);
            font-size: 0.78rem;
            font-family: inherit;
            outline: none;
            transition: border-color 0.15s, box-shadow 0.15s;
            box-sizing: border-box;
            width: 100%;
            height: 30px;
            line-height: 1.2;
        }
        .subst-add-field input:focus,
        .subst-add-field select:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 3px rgba(74, 130, 199, 0.15);
        }
        .subst-add-field input[type="date"] {
            padding: 4px 8px;
        }
        .subst-add-actions {
            flex-direction: row;
            gap: 5px;
            align-items: flex-end;
        }
        .subst-add-actions .btn {
            flex: 1;
            justify-content: center;
            height: 30px;
            padding: 0 10px;
            font-size: 0.72rem;
        }
        .subst-add-actions .btn-outline {
            flex: 0 0 34px;
            padding: 0;
            font-size: 0.85rem;
        }
        .subst-add-error {
            min-height: 14px;
            font-size: 0.68rem;
            font-weight: 600;
            color: var(--danger);
            margin-top: 3px;
        }
        .subst-add-hint {
            font-size: 0.6rem;
            color: var(--text-muted);
            font-style: italic;
            min-height: 10px;
            line-height: 1.2;
            display: none;
        }
        .subst-add-hint.active {
            color: var(--accent);
            font-weight: 600;
            font-style: normal;
            display: block;
        }
        @media (max-width: 1100px) {
            .subst-add-panel .subst-add-grid {
                grid-template-columns: 1fr 1fr;
            }
        }
        @media (max-width: 700px) {
            .subst-add-panel .subst-add-grid {
                grid-template-columns: 1fr;
            }
        }

        .subst-dashboard {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 10px;
            flex-shrink: 0;
        }

        .subst-view-toggle {
            display: inline-flex;
            gap: 2px;
            padding: 3px;
            background: var(--table-bg);
            border: 1.5px solid var(--border);
            border-radius: 30px;
        }
        .subst-view-btn {
            padding: 5px 12px;
            border-radius: 24px;
            border: none;
            background: transparent;
            color: var(--text-secondary);
            font-weight: 600;
            font-size: 0.76rem;
            font-family: inherit;
            cursor: pointer;
            transition: all 0.2s;
            white-space: nowrap;
        }
        .subst-view-btn:hover:not(.active) {
            background: var(--accent-light);
            color: var(--accent);
        }
        .subst-view-btn.active {
            background: var(--accent);
            color: white;
            box-shadow: 0 2px 6px rgba(26, 58, 107, 0.25);
        }
        body.dark-theme .subst-view-btn.active {
            box-shadow: 0 2px 6px rgba(74, 130, 199, 0.3);
        }
        .subst-controls-group {
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }

        .subst-layout {
            display: flex;
            gap: 12px;
            flex: 1 1 auto;
            min-height: 0;
            overflow: hidden;
        }
        .subst-layout .subst-body {
            flex: 1 1 auto;
            min-width: 0;
        }
        .subst-layout .subst-calendar {
            flex: 0 0 300px;
            max-width: 300px;
            display: flex;
            flex-direction: column;
            background: var(--table-bg);
            border: 1px solid var(--border);
            border-radius: 12px;
            box-shadow: var(--shadow-table);
            overflow: hidden;
            flex-shrink: 0;
        }

        .subst-calendar-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 10px 12px;
            background: var(--th-bg);
            color: white;
            font-weight: 700;
            font-size: 0.85rem;
        }
        .subst-cal-nav {
            background: rgba(255,255,255,0.15);
            border: none;
            color: white;
            width: 26px;
            height: 26px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 0.85rem;
            font-family: inherit;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: background 0.15s;
        }
        .subst-cal-nav:hover {
            background: rgba(255,255,255,0.3);
        }
        .subst-cal-month {
            flex: 1;
            text-align: center;
            letter-spacing: 0.3px;
        }
        .subst-cal-weekdays {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            padding: 6px 8px 4px;
            background: var(--accent-light);
            border-bottom: 1px solid var(--border);
        }
        .subst-cal-weekdays span {
            text-align: center;
            font-size: 0.68rem;
            font-weight: 700;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.3px;
        }
        .subst-cal-grid {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 2px;
            padding: 6px 8px 10px;
            overflow-y: auto;
            flex: 1 1 auto;
            min-height: 0;
        }
        .subst-cal-day {
            position: relative;
            aspect-ratio: 1 / 1;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            border-radius: 8px;
            font-size: 0.8rem;
            font-weight: 600;
            color: var(--text-primary);
            cursor: pointer;
            transition: all 0.15s;
            border: 1.5px solid transparent;
            user-select: none;
            background: var(--table-row-even);
            padding: 2px;
            min-height: 38px;
        }
        .subst-cal-day:hover {
            background: var(--accent-lighter);
            border-color: var(--accent);
        }
        .subst-cal-day.other-month {
            opacity: 0.3;
            cursor: default;
            pointer-events: none;
            background: transparent;
        }
        .subst-cal-day.weekend {
            color: var(--text-muted);
        }
        .subst-cal-day.has-substitutions {
            background: var(--substitute-bg);
            border-color: var(--substitute-border);
            color: var(--substitute-text);
        }
        .subst-cal-day.has-substitutions:hover {
            background: var(--substitute-bg-hover);
        }
        .subst-cal-day.today {
            box-shadow: inset 0 0 0 2px var(--accent);
        }
        .subst-cal-day.selected {
            background: var(--accent);
            color: white;
            border-color: var(--accent);
            box-shadow: 0 2px 8px rgba(26, 58, 107, 0.4);
        }
        .subst-cal-day .cal-count {
            font-size: 0.58rem;
            font-weight: 700;
            line-height: 1;
            margin-top: 1px;
            padding: 1px 4px;
            border-radius: 8px;
            background: var(--substitute-badge-bg);
            color: var(--substitute-badge-text);
            min-width: 14px;
            text-align: center;
        }
        .subst-cal-day.selected .cal-count {
            background: rgba(255,255,255,0.3);
            color: white;
        }

        .subst-cal-legend {
            display: flex;
            flex-wrap: wrap;
            gap: 8px 12px;
            padding: 8px 12px;
            border-top: 1px solid var(--border);
            font-size: 0.66rem;
            color: var(--text-secondary);
            background: var(--accent-lighter);
        }
        .subst-cal-legend .dot {
            display: inline-block;
            width: 8px;
            height: 8px;
            border-radius: 50%;
            margin-right: 4px;
            vertical-align: middle;
        }
        .subst-cal-legend .dot-today { background: var(--accent); }
        .subst-cal-legend .dot-has { background: var(--substitute-badge-bg); }
        .subst-cal-legend .dot-selected { background: var(--accent); box-shadow: 0 0 0 2px rgba(26,58,107,0.3); }

        .dict-body { max-height: 400px; overflow-y: auto; padding-right: 4px; }
        .dict-content { display: none; }
        .dict-content.active { display: block; }
        .dict-hint {
            font-size: 0.72rem;
            color: var(--text-muted);
            font-style: italic;
            padding: 8px 10px;
            background: var(--accent-lighter);
            border-radius: 8px;
            margin-bottom: 10px;
            line-height: 1.4;
        }
        .dict-add-row {
            display: flex;
            gap: 6px;
            margin-bottom: 10px;
        }
        .dict-add-row input {
            flex: 1;
            padding: 7px 10px;
            border-radius: 8px;
            border: 1.5px solid var(--border);
            background: var(--table-bg);
            color: var(--text-primary);
            font-size: 0.82rem;
            font-family: inherit;
            outline: none;
        }
        .dict-add-row input:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 3px rgba(74, 130, 199, 0.15);
        }
        .dict-list {
            display: flex;
            flex-direction: column;
            gap: 4px;
        }
        .dict-item {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 7px 10px;
            border-radius: 8px;
            background: var(--table-bg);
            border: 1px solid var(--border);
            font-size: 0.8rem;
            transition: background 0.15s;
        }
        .dict-item:hover { background: var(--table-row-hover); }
        .dict-item.auto { border-left: 3px solid var(--accent); }
        .dict-item.manual { border-left: 3px solid var(--success); }
        .dict-item .dict-name { flex: 1; color: var(--text-primary); font-weight: 500; }
        .dict-item .dict-badge {
            font-size: 0.6rem;
            padding: 1px 6px;
            border-radius: 8px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.3px;
        }
        .dict-item.auto .dict-badge { background: var(--accent-light); color: var(--accent); }
        .dict-item.manual .dict-badge { background: #e5f5e8; color: var(--success); }
        body.dark-theme .dict-item.manual .dict-badge { background: #1e3a26; color: var(--success); }
        .dict-item .icon-btn { font-size: 0.85rem; }
        .dict-empty {
            text-align: center;
            padding: 20px;
            color: var(--text-muted);
            font-style: italic;
            font-size: 0.8rem;
        }

        @media (max-width: 1100px) {
            .subst-layout {
                flex-direction: column;
            }
            .subst-layout .subst-calendar {
                flex: 0 0 auto;
                max-width: 100%;
                max-height: 320px;
            }
            .subst-cal-grid {
                max-height: 220px;
            }
        }

        .constructor-wrap {
            display: flex;
            flex-direction: column;
            gap: 10px;
            height: 100%;
            min-height: 0;
        }
        .constructor-toolbar {
            display: flex;
            gap: 6px;
            flex-wrap: wrap;
            align-items: center;
            padding: 10px 14px;
            background: var(--accent-lighter);
            border: 1px solid var(--border);
            border-radius: 10px;
            flex-shrink: 0;
        }
        .constructor-layout {
            display: flex;
            gap: 12px;
            flex: 1 1 auto;
            min-height: 0;
            overflow: hidden;
        }
        .constructor-nav {
            flex: 0 0 200px;
            display: flex;
            flex-direction: column;
            gap: 3px;
            padding: 8px;
            background: var(--accent-lighter);
            border: 1px solid var(--border);
            border-radius: 10px;
            overflow-y: auto;
        }
        .constr-nav-btn {
            padding: 8px 12px;
            border-radius: 8px;
            border: none;
            background: transparent;
            color: var(--text-secondary);
            font-weight: 600;
            font-size: 0.78rem;
            font-family: inherit;
            text-align: left;
            cursor: pointer;
            transition: all 0.15s;
            white-space: nowrap;
        }
        .constr-nav-btn:hover:not(.active) {
            background: var(--accent-light);
            color: var(--accent);
        }
        .constr-nav-btn.active {
            background: var(--accent);
            color: white;
        }
        .constructor-content {
            flex: 1 1 auto;
            min-width: 0;
            overflow-y: auto;
            padding-right: 6px;
        }
        .constr-section {
            display: none;
            animation: constrFade 0.2s ease;
        }
        .constr-section.active { display: block; }
        @keyframes constrFade {
            from { opacity: 0; transform: translateY(6px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .constr-section h3 {
            font-size: 1rem;
            color: var(--text-primary);
            margin-bottom: 12px;
            padding-bottom: 8px;
            border-bottom: 1.5px solid var(--border);
        }
        .constr-section h4 {
            font-size: 0.85rem;
            color: var(--text-primary);
        }
        .constr-field {
            display: flex;
            flex-direction: column;
            gap: 4px;
            margin-bottom: 12px;
        }
        .constr-field label {
            font-size: 0.72rem;
            font-weight: 600;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.3px;
        }
        .constr-field input[type="text"],
        .constr-field textarea,
        .constr-field select {
            width: 100%;
            padding: 8px 10px;
            border-radius: 8px;
            border: 1.5px solid var(--border);
            background: var(--table-bg);
            color: var(--text-primary);
            font-size: 0.82rem;
            font-family: inherit;
            outline: none;
            transition: border-color 0.15s, box-shadow 0.15s;
            box-sizing: border-box;
        }
        .constr-field input[type="text"]:focus,
        .constr-field textarea:focus,
        .constr-field select:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 3px rgba(74, 130, 199, 0.15);
        }
        .constr-field textarea {
            resize: vertical;
            min-height: 70px;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 0.78rem;
            line-height: 1.4;
        }
        .constr-field-row {
            display: flex;
            flex-wrap: wrap;
            gap: 10px 18px;
            margin-bottom: 12px;
            align-items: center;
        }
        .constr-check {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-size: 0.78rem;
            color: var(--text-primary);
            font-weight: 500;
            cursor: pointer;
            user-select: none;
        }
        .constr-check input[type="checkbox"] {
            cursor: pointer;
            accent-color: var(--accent);
            width: 15px;
            height: 15px;
        }
        .constr-color-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 10px;
        }
        .constr-color-row {
            display: flex;
            gap: 6px;
            align-items: center;
        }
        .constr-color-row input[type="color"] {
            width: 44px;
            height: 34px;
            padding: 2px;
            border: 1.5px solid var(--border);
            border-radius: 6px;
            background: var(--table-bg);
            cursor: pointer;
            flex-shrink: 0;
        }
        .constr-color-row input[type="text"] {
            flex: 1;
            font-family: 'Consolas', monospace;
            font-size: 0.75rem;
            text-transform: lowercase;
        }
        .constr-field input[type="range"] {
            width: 100%;
            accent-color: var(--accent);
            cursor: pointer;
        }
        .constr-hint {
            font-size: 0.72rem;
            color: var(--text-muted);
            font-style: italic;
            margin-bottom: 10px;
            line-height: 1.4;
        }
        .constr-tabs-list {
            display: flex;
            flex-direction: column;
            gap: 6px;
        }
        .constr-tab-row {
            display: grid;
            grid-template-columns: 40px 1fr 110px 100px;
            gap: 8px;
            align-items: center;
            padding: 8px 10px;
            background: var(--table-bg);
            border: 1px solid var(--border);
            border-radius: 8px;
            font-size: 0.8rem;
        }
        .constr-tab-row .constr-tab-idx {
            text-align: center;
            font-weight: 700;
            color: var(--text-muted);
            font-size: 0.72rem;
        }
        .constr-tab-row input[type="text"] {
            padding: 5px 8px;
            font-size: 0.78rem;
            border-radius: 6px;
            border: 1.5px solid var(--border);
            background: var(--table-bg);
            color: var(--text-primary);
            font-family: inherit;
            outline: none;
        }
        .constr-tab-row input[type="text"]:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 2px rgba(74, 130, 199, 0.15);
        }
        .constr-tab-row .constr-tab-actions {
            display: flex;
            gap: 2px;
            justify-content: flex-end;
        }
        .constr-tab-row .constr-tab-actions .icon-btn {
            font-size: 0.8rem;
        }
        .constr-presets {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
        }
        .constr-preset-btn {
            padding: 8px 14px;
            border-radius: 30px;
            border: 1.5px solid var(--border);
            background: var(--table-bg);
            color: var(--text-primary);
            font-weight: 600;
            font-size: 0.76rem;
            font-family: inherit;
            cursor: pointer;
            transition: all 0.15s;
        }
        .constr-preset-btn:hover {
            border-color: var(--accent);
            background: var(--accent-light);
            color: var(--accent);
            transform: translateY(-1px);
        }
        .constr-status {
            padding: 8px 12px;
            font-size: 0.75rem;
            font-weight: 600;
            border-radius: 8px;
            background: var(--accent-lighter);
            color: var(--text-secondary);
            min-height: 20px;
            flex-shrink: 0;
            text-align: center;
        }
        .constr-status.success {
            background: #e5f5e8;
            color: var(--success);
        }
        body.dark-theme .constr-status.success {
            background: #1e3a26;
        }
        .constr-status.error {
            background: #ffe5ea;
            color: var(--danger);
        }
        body.dark-theme .constr-status.error {
            background: #3a1a22;
        }

        .custom-tab-panel {
            display: none;
            flex: 1 1 auto;
            min-height: 0;
            overflow-y: auto;
            padding: 8px;
        }
        body.mode-custom .custom-tab-panel.active {
            display: block;
        }
        body.mode-custom #tableWrapper,
        body.mode-custom #dayFilter,
        body.mode-custom .filter-bar,
        body.mode-custom #substitutionsPanel,
        body.mode-custom #scheduleHint { display: none !important; }

        .custom-content-box {
            padding: 20px 24px;
            background: var(--table-bg);
            border: 1px solid var(--border);
            border-radius: 12px;
            font-size: 0.85rem;
            line-height: 1.5;
            color: var(--text-primary);
            box-shadow: var(--shadow-table);
        }
        .custom-content-box h1,
        .custom-content-box h2,
        .custom-content-box h3 {
            color: var(--text-primary);
            margin: 12px 0 8px;
        }
        .custom-content-box p { margin-bottom: 10px; }
        .custom-content-box ul, .custom-content-box ol { padding-left: 20px; margin-bottom: 10px; }

        @media (max-width: 800px) {
            .constructor-layout { flex-direction: column; }
            .constructor-nav {
                flex: 0 0 auto;
                flex-direction: row;
                overflow-x: auto;
                padding: 6px;
            }
            .constr-nav-btn { white-space: nowrap; }
        }

        .update-wrap {
            display: flex;
            flex-direction: column;
            gap: 12px;
            height: 100%;
            min-height: 0;
        }
        .update-toolbar {
            display: flex;
            gap: 6px;
            flex-wrap: wrap;
            align-items: center;
            padding: 10px 14px;
            background: var(--accent-lighter);
            border: 1px solid var(--border);
            border-radius: 10px;
            flex-shrink: 0;
        }
        .update-grid {
            display: grid;
            grid-template-columns: 1.4fr 1fr;
            gap: 12px;
            flex: 1 1 auto;
            min-height: 0;
            overflow-y: auto;
            padding-right: 4px;
        }
        @media (max-width: 900px) {
            .update-grid { grid-template-columns: 1fr; }
        }
        .update-side {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }
        .update-card {
            background: var(--table-bg);
            border: 1px solid var(--border);
            border-radius: 12px;
            overflow: hidden;
            box-shadow: var(--shadow-table);
            display: flex;
            flex-direction: column;
        }
        .update-card-header {
            padding: 10px 14px;
            background: var(--th-bg);
            color: white;
            font-weight: 700;
            font-size: 0.85rem;
            letter-spacing: 0.3px;
        }
        .update-card-body {
            padding: 12px 14px;
            flex: 1;
        }
        .update-hint {
            font-size: 0.74rem;
            color: var(--text-secondary);
            line-height: 1.5;
            margin-bottom: 12px;
        }
        .update-hint strong { color: var(--danger); }

        .update-drop-zone {
            border: 2px dashed var(--border);
            border-radius: 10px;
            padding: 24px 16px;
            text-align: center;
            transition: all 0.2s ease;
            background: var(--accent-lighter);
            cursor: pointer;
            margin-bottom: 12px;
        }
        .update-drop-zone:hover {
            border-color: var(--accent);
            background: var(--accent-light);
        }
        .update-drop-zone.dragover {
            border-color: var(--accent);
            background: var(--accent-light);
            transform: scale(1.01);
            box-shadow: 0 4px 12px rgba(26, 58, 107, 0.15);
        }
        .update-drop-zone.has-file {
            border-style: solid;
            border-color: var(--success);
            background: #e5f5e8;
        }
        body.dark-theme .update-drop-zone.has-file {
            background: #1e3a26;
        }
        .update-drop-icon {
            font-size: 2.5rem;
            margin-bottom: 8px;
            opacity: 0.7;
        }
        .update-drop-text {
            font-size: 0.82rem;
            color: var(--text-secondary);
            line-height: 1.5;
        }
        .update-drop-text strong { color: var(--text-primary); }
        .update-drop-btn {
            background: none;
            border: none;
            color: var(--accent);
            text-decoration: underline;
            font-weight: 600;
            font-size: inherit;
            cursor: pointer;
            font-family: inherit;
            padding: 0;
            margin: 0;
        }
        .update-drop-btn:hover { color: var(--accent-hover); }

        .update-file-info {
            padding: 10px 12px;
            background: var(--accent-lighter);
            border-radius: 8px;
            border: 1px solid var(--border);
            margin-bottom: 12px;
            font-size: 0.78rem;
        }
        .update-file-row {
            display: flex;
            justify-content: space-between;
            gap: 12px;
            padding: 3px 0;
            border-bottom: 1px dashed var(--border-light);
        }
        .update-file-row:last-child { border-bottom: none; }
        .update-file-label { color: var(--text-secondary); }
        .update-file-info strong { color: var(--text-primary); font-weight: 700; }

        .update-checks {
            padding: 10px 12px;
            background: var(--accent-lighter);
            border-radius: 8px;
            border: 1px solid var(--border);
            margin-bottom: 12px;
        }
        .update-check-item {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 4px 0;
            font-size: 0.76rem;
            color: var(--text-secondary);
            transition: color 0.15s;
        }
        .update-check-item.ok { color: var(--success); }
        .update-check-item.fail { color: var(--danger); }
        .update-check-icon {
            font-size: 0.9rem;
            width: 18px;
            display: inline-block;
            text-align: center;
        }

        .update-error {
            color: var(--danger);
            font-size: 0.76rem;
            font-weight: 600;
            min-height: 16px;
            margin-bottom: 8px;
        }
        .update-actions {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }

        .update-info-row {
            display: flex;
            justify-content: space-between;
            gap: 8px;
            padding: 5px 0;
            border-bottom: 1px solid var(--border-light);
            font-size: 0.78rem;
            color: var(--text-secondary);
        }
        .update-info-row:last-child { border-bottom: none; }
        .update-info-row strong {
            color: var(--text-primary);
            font-weight: 700;
            text-align: right;
        }

        .update-backup-info {
            font-size: 0.74rem;
            color: var(--text-secondary);
            padding: 8px 10px;
            background: var(--accent-lighter);
            border-radius: 6px;
            border: 1px solid var(--border);
            margin-bottom: 4px;
        }
        .update-backup-info strong {
            color: var(--text-primary);
            font-weight: 700;
        }

        .update-status {
            padding: 8px 12px;
            font-size: 0.75rem;
            font-weight: 600;
            border-radius: 8px;
            background: var(--accent-lighter);
            color: var(--text-secondary);
            min-height: 20px;
            flex-shrink: 0;
            text-align: center;
        }
        .update-status.success {
            background: #e5f5e8;
            color: var(--success);
        }
        body.dark-theme .update-status.success {
            background: #1e3a26;
        }
        .update-status.error {
            background: #ffe5ea;
            color: var(--danger);
        }
        body.dark-theme .update-status.error {
            background: #3a1a22;
        }
        .update-status.warning {
            background: #fff3cd;
            color: var(--warning);
        }

        /* ==================== МОБИЛЬНАЯ ВЕРСИЯ ==================== */
        .mobile-toggle {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            padding: 7px 14px;
            border-radius: 30px;
            border: 1.5px solid var(--accent);
            background: transparent;
            color: var(--accent);
            font-weight: 600;
            font-size: 0.8rem;
            font-family: inherit;
            cursor: pointer;
            transition: all 0.2s;
            box-shadow: var(--shadow-btn);
            white-space: nowrap;
            flex-shrink: 0;
        }
        .mobile-toggle:hover {
            background: var(--accent-light);
            transform: translateY(-1px);
        }
        .mobile-toggle.active {
            background: var(--accent);
            color: white;
            border-color: var(--accent);
        }
        .mobile-toggle .mobile-toggle-icon {
            font-size: 0.95rem;
            line-height: 1;
        }

        @media (max-width: 1200px) {
            .mobile-toggle .mobile-toggle-label { display: none; }
            .mobile-toggle { padding: 7px 10px; }
        }

        /* ==================== МОБИЛЬНЫЙ РЕЖИМ ==================== */
        /* Снимаем блокировку прокрутки — критично для телефона */
        html.mobile-view,
        body.mobile-view {
            overflow-y: auto !important;
            overflow-x: hidden !important;
            height: auto !important;
            min-height: 100% !important;
            max-height: none !important;
        }

        body.mobile-view {
            padding: 0;
            align-items: stretch;
            display: block;
        }

        body.mobile-view .animated-bg {
            display: none !important;
        }

        body.mobile-view .app-container {
            padding: 10px 10px 30px;
            gap: 10px;
            border-radius: 0;
            max-height: none !important;
            height: auto !important;
            min-height: 100vh;
            box-shadow: none;
            border: none;
            backdrop-filter: none;
            background: var(--bg-page);
            overflow: visible;
        }

        body.mobile-view .school-header {
            padding-bottom: 10px;
            gap: 8px;
            flex-direction: column;
            align-items: stretch;
        }

        body.mobile-view .school-title {
            text-align: center;
        }

        body.mobile-view .school-title h1 {
            font-size: 1.05rem;
            line-height: 1.25;
        }

        body.mobile-view .school-title p {
            font-size: 0.7rem;
            line-height: 1.3;
            margin-top: 3px;
        }

        body.mobile-view .school-title .subtitle {
            font-size: 0.68rem;
            margin-top: 2px;
        }

        body.mobile-view .toolbar {
            gap: 6px;
            width: 100%;
            justify-content: center;
            flex-wrap: wrap;
        }

        body.mobile-view .schedule-tabs {
            padding: 4px;
            gap: 3px;
            margin-right: 0;
            width: 100%;
            justify-content: space-between;
            overflow-x: auto;
            scrollbar-width: none;
        }
        body.mobile-view .schedule-tabs::-webkit-scrollbar { display: none; }

        body.mobile-view .schedule-tab {
            padding: 8px 10px;
            font-size: 0.7rem;
            flex: 1 1 auto;
            min-width: 60px;
            justify-content: center;
            gap: 3px;
        }

        body.mobile-view .schedule-tab .tab-label {
            display: none;
        }

        body.mobile-view .schedule-tab .tab-icon {
            font-size: 1.15rem;
        }

        body.mobile-view .btn {
            padding: 7px 11px;
            font-size: 0.72rem;
            gap: 4px;
            min-height: 36px;
        }

        body.mobile-view .btn-sm {
            padding: 6px 10px;
            font-size: 0.68rem;
            min-height: 34px;
        }

        body.mobile-view .edit-indicator,
        body.mobile-view .substitute-indicator,
        body.mobile-view .admin-mode-indicator {
            padding: 5px 10px;
            font-size: 0.68rem;
        }

        body.mobile-view .theme-toggle {
            display: none;
        }

        body.mobile-view .schedule-hint {
            padding: 8px 10px;
            font-size: 0.72rem;
            border-radius: 10px;
        }

        body.mobile-view .filter-bar {
            padding: 10px;
            gap: 8px;
            flex-direction: column;
            align-items: stretch;
            border-radius: 10px;
        }

        body.mobile-view .filter-bar .filter-label {
            font-size: 0.68rem;
            margin-bottom: -2px;
        }

        body.mobile-view .search-wrapper {
            flex: 1 1 auto;
            min-width: 0;
            width: 100%;
        }

        body.mobile-view .search-wrapper input {
            padding: 10px 12px 10px 34px;
            font-size: 0.85rem;
            min-height: 42px;
            border-radius: 10px;
        }

        body.mobile-view .search-wrapper .search-icon {
            left: 12px;
            font-size: 1rem;
        }

        body.mobile-view .teacher-select {
            flex: 1 1 auto;
            min-width: 0;
            width: 100%;
            padding: 10px 12px;
            font-size: 0.85rem;
            min-height: 42px;
            border-radius: 10px;
        }

        body.mobile-view .substitute-toggle {
            padding: 10px 12px;
            font-size: 0.78rem;
            justify-content: center;
            min-height: 42px;
            border-radius: 10px;
            width: 100%;
        }

        body.mobile-view .substitute-toggle input {
            width: 18px;
            height: 18px;
        }

        body.mobile-view .filter-count {
            font-size: 0.72rem;
            text-align: center;
        }

        body.mobile-view .day-filter {
            gap: 5px;
            justify-content: stretch;
            flex-wrap: wrap;
        }

        body.mobile-view .day-btn {
            padding: 8px 6px;
            font-size: 0.75rem;
            flex: 1 1 calc(16.666% - 5px);
            min-width: 0;
            text-align: center;
            min-height: 38px;
            border-radius: 10px;
        }

        body.mobile-view .table-wrapper {
            border-radius: 10px;
            overflow: visible;
            background: transparent;
            border: none;
            box-shadow: none;
            flex: none;
            min-height: 0;
        }

        body.mobile-view #scheduleTable {
            min-width: 0;
            width: 100%;
            display: block;
            font-size: 0.85rem;
        }

        body.mobile-view #scheduleTable thead {
            display: none;
        }

        body.mobile-view #scheduleTable tbody {
            display: block;
        }

        body.mobile-view #scheduleTable tr {
            display: block;
            margin-bottom: 12px;
            background: var(--table-bg);
            border-radius: 12px;
            border: 1px solid var(--border);
            box-shadow: 0 2px 6px rgba(0,0,0,0.05);
            overflow: hidden;
        }

        body.mobile-view #scheduleTable tr.no-data-row {
            box-shadow: none;
            background: transparent;
            border: none;
        }

        body.mobile-view #scheduleTable tr.no-data-row td {
            padding: 30px 16px;
            text-align: center;
            background: var(--table-bg) !important;
            border-radius: 12px;
            border: 1px solid var(--border);
        }

        body.mobile-view #scheduleTable td {
            display: block;
            border: none;
            padding: 0;
            background: transparent !important;
            min-width: 0;
        }

        body.mobile-view #scheduleTable td.teacher-name {
            position: static;
            background: var(--th-bg) !important;
            color: white !important;
            font-size: 0.95rem;
            font-weight: 700;
            padding: 10px 12px;
            border-radius: 12px 12px 0 0;
            border-right: none;
            text-align: left;
            letter-spacing: 0.2px;
        }

        body.mobile-view #scheduleTable td.room-cell,
        body.mobile-view #scheduleTable td.class-cell {
            position: static;
            display: inline-block;
            background: var(--accent-lighter) !important;
            color: var(--text-secondary) !important;
            font-size: 0.72rem;
            padding: 6px 10px;
            text-align: left;
            border-radius: 0;
            width: 50%;
            box-sizing: border-box;
            border-bottom: 1px solid var(--border-light);
        }
        body.mobile-view #scheduleTable td.room-cell { border-right: 1px solid var(--border-light); }
        body.mobile-view #scheduleTable td.room-cell::before { content: "🏫 Каб.: "; opacity: 0.7; }
        body.mobile-view #scheduleTable td.class-cell::before { content: "👥 Кл.: "; opacity: 0.7; }

        body.mobile-view #scheduleTable td.room-cell:empty,
        body.mobile-view #scheduleTable td.class-cell:empty {
            display: none;
        }

        body.mobile-view #scheduleTable td.lesson-cell {
            position: relative;
            display: inline-block;
            width: calc(50% - 8px);
            margin: 4px;
            padding: 8px 9px 8px 26px;
            background: var(--table-row-even) !important;
            border-radius: 8px;
            border: 1px solid var(--border-light);
            font-size: 0.75rem;
            line-height: 1.25;
            vertical-align: top;
            min-height: 44px;
            box-sizing: border-box;
            word-break: break-word;
            color: var(--text-primary);
        }

        body.mobile-view #scheduleTable td.lesson-cell::before {
            content: attr(data-lesson-num);
            position: absolute;
            top: 4px;
            left: 5px;
            font-size: 0.62rem;
            font-weight: 800;
            color: var(--accent);
            background: var(--accent-light);
            width: 16px;
            height: 16px;
            line-height: 16px;
            text-align: center;
            border-radius: 50%;
        }

        body.mobile-view #scheduleTable td.lesson-cell.empty-lesson {
            background: transparent !important;
            border-style: dashed;
            border-color: var(--border-light);
            opacity: 0.5;
            min-height: 30px;
        }
        body.mobile-view #scheduleTable td.lesson-cell.empty-lesson::before {
            background: transparent;
            color: var(--text-muted);
            opacity: 0.6;
        }

        body.mobile-view #scheduleTable td.lesson-cell.has-substitute {
            background: var(--substitute-bg) !important;
            border-color: var(--substitute-border);
            border-width: 1.5px;
        }

        body.mobile-view #scheduleTable td.lesson-cell.has-substitute .lesson-text {
            color: var(--substitute-text);
            font-weight: 700;
        }

        body.mobile-view #scheduleTable td.lesson-cell .substitute-note {
            display: block;
            font-size: 0.68rem;
            margin-top: 3px;
            color: var(--substitute-text);
            font-weight: 600;
            font-style: italic;
            line-height: 1.2;
        }

        body.mobile-view #scheduleTable td.lesson-cell .substitute-icon {
            top: 3px;
            right: 4px;
            font-size: 0.7rem;
        }

        body.mobile-view #scheduleTable td.lesson-cell.editing {
            width: calc(100% - 8px);
            padding: 8px;
        }

        body.mobile-view .cell-editor {
            min-width: 0;
            width: 100%;
        }

        body.mobile-view .cell-editor input[type="text"] {
            font-size: 0.85rem;
            padding: 8px 10px;
            min-height: 38px;
        }

        body.mobile-view .free-teachers-dropdown {
            max-height: 260px;
        }

        body.mobile-view .free-teachers-dropdown .dropdown-item {
            padding: 10px 12px;
            font-size: 0.82rem;
        }

        body.mobile-view .subst-add-panel {
            padding: 10px;
            border-radius: 10px;
        }

        body.mobile-view .subst-add-title {
            font-size: 0.85rem;
            margin-bottom: 8px;
        }

        body.mobile-view .subst-add-grid {
            grid-template-columns: 1fr !important;
            gap: 8px;
        }

        body.mobile-view .subst-add-field label {
            font-size: 0.68rem;
            margin-bottom: 2px;
        }

        body.mobile-view .subst-add-field input,
        body.mobile-view .subst-add-field select {
            font-size: 0.85rem;
            height: 40px;
            padding: 8px 10px;
            border-radius: 8px;
        }

        body.mobile-view .subst-add-actions {
            flex-direction: row;
        }
        body.mobile-view .subst-add-actions .btn {
            height: 40px;
            font-size: 0.78rem;
        }
        body.mobile-view .subst-add-actions .btn-outline {
            flex: 0 0 40px;
        }

        body.mobile-view .subst-dashboard {
            grid-template-columns: 1fr 1fr;
            gap: 8px;
        }

        body.mobile-view .subst-stat-card {
            padding: 10px 12px;
            border-radius: 10px;
        }

        body.mobile-view .subst-stat-card .label {
            font-size: 0.65rem;
        }

        body.mobile-view .subst-stat-card .value {
            font-size: 1.35rem;
        }

        body.mobile-view .subst-toolbar {
            padding: 10px;
            gap: 8px;
            flex-direction: column;
            align-items: stretch;
            border-radius: 10px;
        }

        body.mobile-view .subst-toolbar select,
        body.mobile-view .subst-toolbar input[type="text"],
        body.mobile-view .subst-toolbar input[type="month"],
        body.mobile-view .subst-toolbar input[type="date"] {
            padding: 9px 12px;
            font-size: 0.82rem;
            width: 100%;
            min-height: 40px;
            border-radius: 10px;
        }

        body.mobile-view .subst-view-toggle {
            width: 100%;
            justify-content: center;
            padding: 4px;
        }

        body.mobile-view .subst-view-btn {
            flex: 1;
            padding: 8px 10px;
            font-size: 0.78rem;
            text-align: center;
        }

        body.mobile-view .subst-controls-group {
            width: 100%;
            justify-content: center;
            flex-wrap: wrap;
            gap: 6px;
        }

        body.mobile-view .subst-controls-group .btn {
            min-height: 40px;
        }

        body.mobile-view .subst-section {
            border: none;
            background: transparent;
            box-shadow: none;
            border-radius: 0;
        }

        body.mobile-view .subst-section-header {
            border-radius: 10px;
            padding: 10px 12px;
            font-size: 0.85rem;
        }

        body.mobile-view .subst-table {
            display: block;
            font-size: 0.85rem;
        }

        body.mobile-view .subst-table thead {
            display: none;
        }

        body.mobile-view .subst-table tbody {
            display: block;
        }

        body.mobile-view .subst-table tr {
            display: block;
            margin-bottom: 10px;
            background: var(--table-bg);
            border: 1px solid var(--border);
            border-radius: 12px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.05);
            padding: 10px 12px;
            position: relative;
        }

        body.mobile-view .subst-table tr.empty-row {
            padding: 0;
            box-shadow: none;
            background: transparent;
            border: none;
        }

        body.mobile-view .subst-table tr.empty-row td {
            padding: 30px 16px !important;
            text-align: center;
            background: var(--table-bg) !important;
            border-radius: 12px;
            border: 1px solid var(--border);
            display: block;
        }

        body.mobile-view .subst-table td {
            display: block;
            padding: 3px 0;
            border: none;
            background: transparent !important;
            font-size: 0.82rem;
            line-height: 1.35;
            text-align: left !important;
        }

        body.mobile-view .subst-table td:nth-child(1) {
            font-weight: 700;
            color: var(--accent);
            font-size: 0.9rem;
            padding-bottom: 1px;
        }

        body.mobile-view .subst-table td:nth-child(2) {
            display: inline-block;
            color: var(--text-muted);
            font-size: 0.72rem;
            font-weight: 600;
            margin-left: 6px;
            padding: 0;
        }

        body.mobile-view .subst-table td:nth-child(3) {
            margin: 6px 0 8px;
        }

        body.mobile-view .subst-table td:nth-child(4)::before {
            content: "👤 Отсутствует: ";
            color: var(--text-muted);
            font-weight: 600;
            font-size: 0.72rem;
        }
        body.mobile-view .subst-table td:nth-child(5)::before {
            content: "📖 Предмет: ";
            color: var(--text-muted);
            font-weight: 600;
            font-size: 0.72rem;
        }
        body.mobile-view .subst-table td:nth-child(6)::before {
            content: "🔄 Заменяет: ";
            color: var(--substitute-text);
            font-weight: 600;
            font-size: 0.72rem;
        }
        body.mobile-view .subst-table td:nth-child(7)::before {
            content: "📗 Предмет: ";
            color: var(--text-muted);
            font-weight: 600;
            font-size: 0.72rem;
        }

        body.mobile-view .subst-table td.actions {
            position: absolute;
            top: 8px;
            right: 8px;
            width: auto !important;
            padding: 0 !important;
            display: flex !important;
            gap: 4px;
        }

        body.mobile-view .subst-table td.actions .icon-btn {
            font-size: 1rem;
            padding: 5px 7px;
            background: var(--accent-light);
            border-radius: 6px;
        }

        body.mobile-view .subst-table td.actions .icon-btn.danger {
            background: #ffe5ea;
        }

        body.mobile-view .subst-layout {
            flex-direction: column;
            gap: 10px;
        }

        body.mobile-view .subst-layout .subst-calendar {
            max-width: 100%;
            flex: 0 0 auto;
            max-height: none;
        }

        body.mobile-view .subst-cal-day {
            font-size: 0.85rem;
            min-height: 42px;
            border-radius: 10px;
        }

        body.mobile-view .subst-cal-day .cal-count {
            font-size: 0.62rem;
            padding: 1px 5px;
        }

        body.mobile-view .footer-note {
            font-size: 0.72rem;
            padding-top: 10px;
            flex-wrap: wrap;
            justify-content: center;
            gap: 6px;
            line-height: 1.4;
        }

        body.mobile-view .status-badge {
            font-size: 0.66rem;
            padding: 3px 8px;
        }

        body.mobile-view .modal {
            padding: 18px 16px;
            border-radius: 14px;
            max-width: 100%;
            max-height: calc(100vh - 20px);
        }

        body.mobile-view .modal.modal-wide {
            max-width: 100%;
            padding: 12px;
            height: calc(100vh - 16px);
            max-height: calc(100vh - 16px);
        }

        body.mobile-view .modal h2 {
            font-size: 1.05rem;
        }

        body.mobile-view .modal p {
            font-size: 0.8rem;
        }

        body.mobile-view .modal label {
            font-size: 0.8rem;
        }

        body.mobile-view .modal input[type="password"],
        body.mobile-view .modal input[type="text"],
        body.mobile-view .modal input[type="time"],
        body.mobile-view .modal input[type="date"],
        body.mobile-view .modal select {
            padding: 10px 12px;
            font-size: 0.9rem;
            min-height: 42px;
            border-radius: 10px;
        }

        body.mobile-view .modal-actions {
            flex-direction: column-reverse;
            gap: 8px;
        }

        body.mobile-view .modal-actions .btn {
            width: 100%;
            padding: 12px;
            font-size: 0.9rem;
            justify-content: center;
        }

        body.mobile-view #consoleModal .modal.modal-wide {
            padding: 10px;
            height: calc(100vh - 12px);
            max-height: calc(100vh - 12px);
        }

        body.mobile-view .console-header {
            padding-bottom: 8px;
            margin-bottom: 10px;
        }

        body.mobile-view .console-header h2 {
            font-size: 0.95rem;
        }

        body.mobile-view .console-tabs {
            overflow-x: auto;
            padding-bottom: 6px;
            gap: 4px;
            scrollbar-width: none;
        }
        body.mobile-view .console-tabs::-webkit-scrollbar { display: none; }

        body.mobile-view .console-tab {
            padding: 8px 12px;
            font-size: 0.72rem;
            white-space: nowrap;
        }

        body.mobile-view .admin-table {
            display: block;
            font-size: 0.85rem;
        }
        body.mobile-view .admin-table thead { display: none; }
        body.mobile-view .admin-table tbody { display: block; }
        body.mobile-view .admin-table tr {
            display: block;
            margin-bottom: 8px;
            padding: 10px 12px;
            background: var(--table-bg);
            border: 1px solid var(--border);
            border-radius: 10px;
            position: relative;
        }
        body.mobile-view .admin-table td {
            display: block;
            padding: 3px 0;
            border: none;
            background: transparent;
            font-size: 0.82rem;
        }
        body.mobile-view .admin-table td:first-child {
            position: absolute;
            top: 10px;
            right: 12px;
            font-weight: 700;
            color: var(--text-muted);
            font-size: 0.72rem;
            width: auto;
        }
        body.mobile-view .admin-table td:nth-child(2)::before {
            content: "👤 ";
            opacity: 0.7;
        }
        body.mobile-view .admin-table td:nth-child(3)::before {
            content: "🏫 Каб.: ";
            color: var(--text-muted);
            font-size: 0.72rem;
            font-weight: 600;
        }
        body.mobile-view .admin-table td:nth-child(4)::before {
            content: "👥 Кл.: ";
            color: var(--text-muted);
            font-size: 0.72rem;
            font-weight: 600;
        }
        body.mobile-view .admin-table td:nth-child(2),
        body.mobile-view .admin-table td:nth-child(3),
        body.mobile-view .admin-table td:nth-child(4) {
            padding-right: 40px;
        }

        body.mobile-view .admin-table td input {
            font-size: 0.85rem;
            padding: 6px 8px;
            min-height: 34px;
            border-radius: 8px;
        }

        body.mobile-view .admin-table td:last-child {
            padding-top: 8px;
        }
        body.mobile-view .admin-actions {
            justify-content: flex-start;
            gap: 6px;
        }
        body.mobile-view .admin-actions .icon-btn {
            font-size: 1rem;
            padding: 8px 12px;
            background: var(--accent-light);
            border-radius: 8px;
            min-width: 38px;
            text-align: center;
        }
        body.mobile-view .admin-actions .icon-btn.danger {
            background: #ffe5ea;
        }

        body.mobile-view .my-sched-controls {
            padding: 10px;
            gap: 10px;
            flex-direction: column;
        }
        body.mobile-view .my-sched-control {
            flex: 1 1 100%;
            width: 100%;
            min-width: 0;
        }
        body.mobile-view .my-sched-control label {
            font-size: 0.75rem;
        }
        body.mobile-view .my-sched-control > div {
            width: 100%;
        }
        body.mobile-view .my-sched-control > div .btn {
            flex: 1;
        }

        body.mobile-view .my-sched-day-header {
            padding: 10px 12px;
            font-size: 0.9rem;
        }
        body.mobile-view .my-sched-day-header .substitutions-count {
            font-size: 0.7rem;
            padding: 3px 8px;
        }

        body.mobile-view .my-sched-lesson {
            grid-template-columns: 36px 1fr;
            grid-template-areas: "num subject" "num info";
            gap: 4px 10px;
            padding: 10px 12px;
            font-size: 0.85rem;
        }

        body.mobile-view .my-sched-lesson .num {
            grid-area: num;
            align-self: center;
            font-size: 1rem;
        }

        body.mobile-view .my-sched-lesson .subject {
            grid-area: subject;
            font-size: 0.9rem;
        }

        body.mobile-view .my-sched-lesson .time {
            display: none;
        }

        body.mobile-view .my-sched-lesson .class-name,
        body.mobile-view .my-sched-lesson .room {
            grid-area: info;
            display: inline;
            font-size: 0.78rem;
            text-align: left;
        }

        body.mobile-view .my-sched-lesson .class-name::after {
            content: " · ";
        }

        body.mobile-view .my-sched-replacement-row {
            grid-template-columns: 36px 1fr;
            grid-template-areas: "num info" "num who";
            gap: 4px 10px;
            padding: 10px 0;
            font-size: 0.85rem;
        }

        body.mobile-view .my-sched-replacement-row .num {
            grid-area: num;
        }

        body.mobile-view .my-sched-replacement-row .time {
            display: none;
        }

        body.mobile-view .my-sched-replacement-row > div:nth-child(3) {
            grid-area: info;
        }

        body.mobile-view .my-sched-replacement-row > div:nth-child(4) {
            grid-area: who;
            font-size: 0.78rem;
        }

        body.mobile-view .constructor-nav {
            flex-direction: row;
            overflow-x: auto;
            padding: 6px;
            gap: 4px;
            scrollbar-width: none;
        }
        body.mobile-view .constructor-nav::-webkit-scrollbar { display: none; }

        body.mobile-view .constr-nav-btn {
            padding: 8px 12px;
            font-size: 0.72rem;
            white-space: nowrap;
        }

        body.mobile-view .constr-tab-row {
            grid-template-columns: 30px 1fr;
            grid-template-areas: "idx label" "idx icon" "idx actions";
            gap: 6px;
            padding: 10px;
        }
        body.mobile-view .constr-tab-row .constr-tab-idx {
            grid-area: idx;
            align-self: start;
        }
        body.mobile-view .constr-tab-row input[type="text"]:nth-of-type(1) { grid-area: label; }
        body.mobile-view .constr-tab-row input[type="text"]:nth-of-type(2) { grid-area: icon; }
        body.mobile-view .constr-tab-row .constr-tab-actions { grid-area: actions; justify-content: flex-start; }

        body.mobile-view .update-toolbar {
            padding: 10px;
            gap: 8px;
            flex-direction: column;
            align-items: stretch;
        }
        body.mobile-view .update-toolbar .btn {
            width: 100%;
            justify-content: center;
            min-height: 42px;
        }

        body.mobile-view .update-grid {
            grid-template-columns: 1fr;
            gap: 10px;
        }

        body.mobile-view .update-card-header {
            font-size: 0.82rem;
            padding: 10px 12px;
        }

        body.mobile-view .update-card-body {
            padding: 12px;
        }

        body.mobile-view .update-drop-zone {
            padding: 30px 16px;
        }

        body.mobile-view .update-drop-icon {
            font-size: 2rem;
        }

        body.mobile-view .update-actions .btn {
            flex: 1;
            justify-content: center;
        }

        body.mobile-view .dict-add-row {
            flex-direction: column;
        }
        body.mobile-view .dict-add-row .btn {
            width: 100%;
            justify-content: center;
        }

        body.mobile-view #bellsEditor {
            grid-template-columns: 40px 1fr 1fr !important;
            gap: 6px;
        }
        body.mobile-view #bellsEditor input[type="time"] {
            font-size: 0.8rem;
            padding: 8px 6px;
            min-height: 38px;
        }

        @media print {
            body.mobile-view .app-container { background: white; }
        }
    </style>
</head>
<body>
    <div class="animated-bg" id="animatedBg"></div>

    <div class="app-container">
        <div class="school-header">
            <div class="school-title">
                <h1>МОУ «Северная СОШ №2»</h1>
                <p>Белгородского муниципального округа Белгородской области</p>
                <div class="subtitle">Расписание · 2026–2027 учебный год</div>
            </div>
            <div class="toolbar">
                <div class="schedule-tabs" id="scheduleTabs">
                    <button class="schedule-tab active" data-schedule="teachers" title="Расписание учителей">
                        <span class="tab-icon">👨‍🏫</span>
                        <span class="tab-label">Учительское</span>
                    </button>
                    <button class="schedule-tab" data-schedule="classes" title="Расписание по классам">
                        <span class="tab-icon">🎓</span>
                        <span class="tab-label">Детское</span>
                    </button>
                    <button class="schedule-tab" data-schedule="iup" title="Индивидуальные учебные планы">
                        <span class="tab-icon">📚</span>
                        <span class="tab-label">ИУП 5-8</span>
                    </button>
                    <button class="schedule-tab" data-schedule="substitutions" title="Журнал замен учителей">
                        <span class="tab-icon">📌</span>
                        <span class="tab-label">Журнал замен</span>
                    </button>
                </div>

                <div class="edit-indicator" id="editIndicator">
                    <span class="dot"></span>
                    Режим редактирования
                </div>
                <div class="substitute-indicator" id="substituteIndicator">
                    📌 Замен: <span id="substituteIndicatorCount">0</span>
                </div>
                <button class="btn btn-outline btn-sm" id="resetBtn" style="display:none;">↺ Сбросить</button>
                <button class="btn btn-success btn-sm" id="saveBtn" style="display:none;">💾 Сохранить</button>
                <button class="btn btn-primary btn-sm" id="editToggleBtn">✎ Редактировать</button>

                <div class="admin-mode-indicator" id="adminIndicator">
                    <span class="admin-dot"></span>
                    Администратор
                </div>
                <button class="btn btn-logout btn-sm" id="logoutAdminBtn" style="display:none;">🚪 Выйти</button>

                <button class="btn btn-outline btn-sm" id="myScheduleBtn" title="Моё расписание">👤 Моё расписание</button>

                <div class="console-btn-wrapper">
                    <button class="btn btn-outline btn-sm" id="consoleBtn" title="Консоль администратора">⚙ Консоль</button>
                    <span class="console-badge hidden" id="consoleBadge">0</span>
                </div>

                <button class="mobile-toggle" id="mobileToggle" title="Переключить мобильную версию">
                    <span class="mobile-toggle-icon" id="mobileToggleIcon">📱</span>
                    <span class="mobile-toggle-label" id="mobileToggleLabel">Мобильная</span>
                </button>

                <button class="theme-toggle" id="themeToggle" title="Переключить тему" aria-label="Переключить тему">
                    <span class="knob">
                        <span class="icon-sun">☀️</span>
                        <span class="icon-moon">🌙</span>
                    </span>
                </button>
            </div>
        </div>

        <div class="schedule-hint" id="scheduleHint">
            <span class="hint-icon">ℹ️</span>
            <span id="scheduleHintText"><strong>Раздел администратора</strong>: редактирование, замены, консоль, история изменений.</span>
        </div>

        <div class="filter-bar">
            <span class="filter-label">🔍 Поиск</span>
            <div class="search-wrapper">
                <span class="search-icon">🔎</span>
                <input type="text" id="teacherSearch" placeholder="Имя учителя..." autocomplete="off">
                <button class="clear-search" id="clearSearch" title="Очистить">✕</button>
            </div>

            <select class="teacher-select" id="teacherSelect" title="Выбрать">
                <option value="">— Все учителя —</option>
            </select>

            <label class="substitute-toggle" title="Показать только ячейки с заменами">
                <input type="checkbox" id="substituteFilter">
                📌 Только замены
            </label>

            <span class="filter-count" id="filterCount"></span>
        </div>

        <div class="day-filter" id="dayFilter">
            <button class="day-btn active" data-day="all">Все дни</button>
            <button class="day-btn" data-day="ПОНЕДЕЛЬНИК">Пн</button>
            <button class="day-btn" data-day="ВТОРНИК">Вт</button>
            <button class="day-btn" data-day="СРЕДА">Ср</button>
            <button class="day-btn" data-day="ЧЕТВЕРГ">Чт</button>
            <button class="day-btn" data-day="ПЯТНИЦА">Пт</button>
        </div>

        <div class="table-wrapper" id="tableWrapper">
            <table id="scheduleTable">
                <thead id="scheduleThead"></thead>
                <tbody id="scheduleBody"></tbody>
            </table>
        </div>

        <div id="substitutionsPanel">

            <div class="subst-add-panel">
                <div class="subst-add-title">➕ Новая замена</div>
                <div class="subst-add-grid">
                    <div class="subst-add-field">
                        <label for="substDateInput">📅 Дата</label>
                        <input type="date" id="substDateInput">
                        <div class="subst-add-hint" id="substDateHint"></div>
                    </div>

                    <div class="subst-add-field">
                        <label for="substLessonType">🎯 Тип урока</label>
                        <select id="substLessonType">
                            <option value="Стандартный урок">Стандартный урок</option>
                            <option value="Внеурочная деятельность">Внеурочная деятельность</option>
                        </select>
                    </div>

                    <div class="subst-add-field">
                        <label for="substAbsentTeacher">👤 Отсутствующий учитель</label>
                        <select id="substAbsentTeacher">
                            <option value="">— Выберите —</option>
                        </select>
                    </div>

                    <div class="subst-add-field">
                        <label for="substAbsentSubject">📖 Его предмет</label>
                        <select id="substAbsentSubject">
                            <option value="">— Выберите —</option>
                        </select>
                    </div>

                    <div class="subst-add-field">
                        <label for="substSubstituteTeacher">🔄 Заменяющий учитель</label>
                        <select id="substSubstituteTeacher">
                            <option value="">— Выберите —</option>
                        </select>
                    </div>

                    <div class="subst-add-field">
                        <label for="substSubstituteSubject">📗 Его предмет</label>
                        <select id="substSubstituteSubject">
                            <option value="">— Выберите —</option>
                        </select>
                    </div>

                    <div class="subst-add-field subst-add-actions">
                        <button class="btn btn-success btn-sm" id="substAddSubmitBtn">➕ Добавить</button>
                        <button class="btn btn-outline btn-sm" id="substClearFormBtn" title="Очистить форму">↺</button>
                    </div>
                </div>
                <div class="subst-add-error" id="substAddError"></div>
            </div>

            <div class="subst-dashboard">
                <div class="subst-stat-card">
                    <div class="label">Всего замен</div>
                    <div class="value" id="substStatTotal">0</div>
                </div>
                <div class="subst-stat-card">
                    <div class="label">Учителей заменено</div>
                    <div class="value" id="substStatReplaced">0</div>
                </div>
                <div class="subst-stat-card">
                    <div class="label">Учителей заменяло</div>
                    <div class="value" id="substStatSubstitutes">0</div>
                </div>
                <div class="subst-stat-card">
                    <div class="label">За текущий месяц</div>
                    <div class="value subst" id="substStatMonth">0</div>
                </div>
            </div>

            <div class="subst-toolbar">
                <div class="subst-view-toggle" id="substViewToggle">
                    <button class="subst-view-btn active" data-view="month">📅 Месяц</button>
                    <button class="subst-view-btn" data-view="day">📆 День</button>
                </div>

                <span id="substMonthControls" class="subst-controls-group">
                    <span class="filter-label">Период</span>
                    <input type="month" id="substMonthFilter" title="Месяц">
                </span>

                <span id="substDayControls" class="subst-controls-group" style="display:none;">
                    <button class="btn btn-outline btn-sm" id="substPrevDayBtn" title="Предыдущий день">◀</button>
                    <input type="date" id="substDayFilter" title="Выберите день">
                    <button class="btn btn-outline btn-sm" id="substNextDayBtn" title="Следующий день">▶</button>
                    <button class="btn btn-outline btn-sm" id="substTodayBtn" title="Сегодня">Сегодня</button>
                </span>

                <span class="filter-label">🔍 Поиск</span>
                <input type="text" id="substSearch" placeholder="ФИО или предмет..." autocomplete="off">

                <span class="spacer" style="flex:1;"></span>

                <button class="btn btn-outline btn-sm" id="substCalendarToggleBtn" title="Календарь замен">🗓 Календарь</button>
                <button class="btn btn-outline btn-sm" id="substManageDictionariesBtn" title="Справочники">⚙ Справочники</button>
                <button class="btn btn-primary btn-sm" id="substExportExcelBtn">📤 Excel</button>
                <button class="btn btn-outline btn-sm" id="substPrintBtn">🖨 Печать</button>
            </div>

            <div class="subst-toolbar" style="justify-content:space-between;">
                <span class="filter-count" id="substFilterCount"></span>
                <button class="btn btn-outline btn-sm" id="substHistoryToggleBtn">📜 История журнала</button>
            </div>

            <div class="subst-layout">
                <div class="subst-body" id="substBody">
                    <div class="subst-section">
                        <div class="subst-section-header">
                            <span id="substSectionTitle">📋 Журнал замен</span>
                            <span class="count-badge" id="substDetailsCount">0</span>
                        </div>
                        <div>
                            <table class="subst-table" id="substDetailsTable">
                                <thead>
                                    <tr>
                                        <th style="width:90px;">Дата</th>
                                        <th style="width:55px;">День</th>
                                        <th style="width:130px;">Тип урока</th>
                                        <th>Отсутствующий</th>
                                        <th>Его предмет</th>
                                        <th>Заменяющий</th>
                                        <th>Его предмет</th>
                                        <th style="width:80px;">Действия</th>
                                    </tr>
                                </thead>
                                <tbody id="substDetailsBody"></tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <aside class="subst-calendar" id="substCalendar" style="display:none;">
                    <div class="subst-calendar-header">
                        <button class="subst-cal-nav" id="substCalPrevBtn" title="Предыдущий месяц">◀</button>
                        <span class="subst-cal-month" id="substCalMonthLabel">Сентябрь 2026</span>
                        <button class="subst-cal-nav" id="substCalNextBtn" title="Следующий месяц">▶</button>
                    </div>
                    <div class="subst-cal-weekdays">
                        <span>Пн</span><span>Вт</span><span>Ср</span><span>Чт</span><span>Пт</span><span>Сб</span><span>Вс</span>
                    </div>
                    <div class="subst-cal-grid" id="substCalGrid"></div>
                    <div class="subst-cal-legend">
                        <span><span class="dot dot-today"></span>Сегодня</span>
                        <span><span class="dot dot-has"></span>Есть замены</span>
                        <span><span class="dot dot-selected"></span>Выбрано</span>
                    </div>
                </aside>
            </div>

            <div class="subst-history-panel" id="substHistoryPanel">
                <div class="header">
                    <h3>📜 История изменений журнала <span class="status-badge" id="substHistoryCount" style="background: var(--accent);">0</span></h3>
                    <div style="display:flex; gap:6px; flex-wrap:wrap;">
                        <button class="btn btn-outline btn-sm" id="substUndoLastBtn">↶ Отменить последнее</button>
                        <button class="btn btn-outline btn-sm" id="substCloseHistoryBtn">Свернуть</button>
                    </div>
                </div>
                <div class="subst-history-list" id="substHistoryList"></div>
            </div>

        </div>

        <div class="history-panel" id="historyPanel">
            <div class="history-panel-header">
                <h3>📜 История изменений <span class="status-badge" id="historyCount" style="background: var(--accent);">0</span></h3>
                <div style="display:flex; gap:6px; flex-wrap:wrap;">
                    <button class="btn btn-outline btn-sm" id="undoLastBtn">↶ Отменить</button>
                    <button class="btn btn-outline btn-sm" id="toggleHistoryBtn">Свернуть</button>
                </div>
            </div>
            <div class="history-panel-list" id="historyPanelList"></div>
        </div>

        <div class="footer-note">
            <span id="footerHint">Просмотр расписания. Для редактирования нажмите «Редактировать».</span>
            · <span class="status-badge" id="statusBadge">Сохранено локально</span>
            · <button class="btn btn-outline btn-sm" id="showHistoryBtn" style="margin-left:6px;">📜 История</button>
        </div>
    </div>

    <div class="modal-overlay" id="passwordModal">
        <div class="modal modal-narrow">
            <h2 id="passwordModalTitle">🔒 Вход</h2>
            <p id="passwordModalDesc">Введите пароль.</p>
            <label for="passwordInput">Пароль</label>
            <input type="password" id="passwordInput" placeholder="Введите пароль..." autocomplete="off">
            <span class="error-msg" id="passwordError"></span>
            <div class="modal-actions">
                <button class="btn btn-outline" id="cancelPasswordBtn">Отмена</button>
                <button class="btn btn-primary" id="confirmPasswordBtn">Войти</button>
            </div>
        </div>
    </div>

    <div class="modal-overlay" id="consoleModal">
        <div class="modal modal-wide">
            <div class="console-header">
                <h2>🛠 Консоль администратора</h2>
                <button class="btn btn-outline btn-sm" id="closeConsoleBtn">✕ Закрыть</button>
            </div>

            <div class="console-tabs">
                <button class="console-tab active" data-tab="teachers">👥 Учителя</button>
                <button class="console-tab" data-tab="bulk">⚡ Массовые операции</button>
                <button class="console-tab" data-tab="io">💾 Импорт/Экспорт</button>
                <button class="console-tab" data-tab="history">📜 История</button>
                <button class="console-tab locked" data-tab="security" data-requires-password="true">🔐 Безопасность</button>
                <button class="console-tab locked" data-tab="constructor" data-requires-password="true">🎨 Конструктор</button>
                <button class="console-tab locked" data-tab="update" data-requires-password="true">🔄 Обновление сайта</button>
                <button class="console-tab locked" data-tab="danger" data-requires-password="true">⚠ Опасная зона</button>
            </div>

            <div class="console-body">
                <div class="console-tab-content active" data-tab="teachers">
                    <div class="console-toolbar">
                        <button class="btn btn-success btn-sm" id="addTeacherBtn">+ Добавить учителя</button>
                        <span class="spacer"></span>
                        <span class="filter-count" id="adminCount"></span>
                    </div>
                    <table class="admin-table">
                        <thead>
                            <tr>
                                <th style="width:36px;">#</th>
                                <th>Имя учителя</th>
                                <th style="width:70px;">Каб.</th>
                                <th style="width:70px;">Кл.</th>
                                <th style="width:130px;">Действия</th>
                            </tr>
                        </thead>
                        <tbody id="adminTableBody"></tbody>
                    </table>
                </div>

                <div class="console-tab-content" data-tab="bulk">
                    <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px; align-items: start;">
                        <div style="padding:12px; background:var(--accent-lighter); border-radius:10px; border:1px solid var(--border);">
                            <strong style="font-size:0.8rem; color:var(--text-primary);">⚡ Массовое действие</strong>
                            <div style="margin-top:8px;">
                                <label style="font-size:0.7rem; font-weight:600; color:var(--text-secondary); display:block; margin-bottom:4px;">Выберите действие:</label>
                                <select id="bulkAction" class="teacher-select" style="width:100%; margin:0;">
                                    <option value="">— Выберите действие —</option>
                                    <option value="clearTeacher">Очистить расписание учителя</option>
                                    <option value="clearDay">Очистить день у всех</option>
                                    <option value="clearSubstitutes">Убрать все замены</option>
                                    <option value="clearAll">Очистить всё расписание</option>
                                </select>
                            </div>

                            <div id="bulkTeacherPicker" style="display:none; margin-top:8px;">
                                <label style="font-size:0.7rem; font-weight:600; color:var(--text-secondary); display:block; margin-bottom:4px;">Учитель:</label>
                                <select id="bulkTeacherSelect" class="teacher-select" style="width:100%; margin:0;"></select>
                            </div>

                            <div id="bulkDayPicker" style="display:none; margin-top:8px;">
                                <label style="font-size:0.7rem; font-weight:600; color:var(--text-secondary); display:block; margin-bottom:4px;">День недели:</label>
                                <select id="bulkDaySelect" class="teacher-select" style="width:100%; margin:0;">
                                    <option value="ПОНЕДЕЛЬНИК">Понедельник</option>
                                    <option value="ВТОРНИК">Вторник</option>
                                    <option value="СРЕДА">Среда</option>
                                    <option value="ЧЕТВЕРГ">Четверг</option>
                                    <option value="ПЯТНИЦА">Пятница</option>
                                </select>
                            </div>

                            <button class="btn btn-warning btn-sm" id="applyBulkBtn" disabled style="margin-top:10px;">Применить</button>
                        </div>

                        <div style="padding:12px; background:var(--accent-lighter); border-radius:10px; border:1px solid var(--border);">
                            <strong style="font-size:0.8rem; color:var(--text-primary);">📊 Статистика загруженности</strong>
                            <div id="bulkStats" style="margin-top:8px; font-size:0.72rem; color:var(--text-secondary); line-height:1.6;"></div>
                        </div>
                    </div>
                </div>

                <div class="console-tab-content" data-tab="io">
                    <div id="io-tab-grid">
                        <div>
                            <div class="io-block io-excel">
                                <strong>📊 Excel (.xlsx / .xls)</strong>
                                <p>Импорт и экспорт расписания в формате Excel.</p>
                                <div style="display:flex; gap:6px; flex-wrap:wrap;">
                                    <button class="btn btn-success btn-sm" id="importExcelBtn">📥 Импорт</button>
                                    <button class="btn btn-primary btn-sm" id="exportExcelBtn">📤 Расписание</button>
                                    <button class="btn btn-outline btn-sm" id="exportExcelHistoryBtn">📜 История</button>
                                </div>
                                <input type="file" id="importExcelFile" accept=".xlsx,.xls" style="display:none;">

                                <div id="excelImportOptions" style="display:none; margin-top:10px; padding-top:10px; border-top:1px dashed var(--border);">
                                    <div style="display:grid; grid-template-columns: 1fr 1fr; gap:8px;">
                                        <div>
                                            <label style="font-size:0.66rem; font-weight:600; color:var(--text-secondary); display:block; margin-bottom:3px;">Лист Excel:</label>
                                            <select id="excelSheetSelect" class="teacher-select" style="width:100%; margin:0;"></select>
                                        </div>
                                        <div>
                                            <label style="font-size:0.66rem; font-weight:600; color:var(--text-secondary); display:block; margin-bottom:3px;">Строка-заголовок:</label>
                                            <input type="number" id="excelHeaderRow" value="1" min="1" style="width:100%; padding:6px 10px; border-radius:6px; border:1.5px solid var(--border); background:var(--table-bg); color:var(--text-primary); font-family:inherit; font-size:0.72rem; box-sizing:border-box;">
                                        </div>
                                    </div>
                                    <div id="excelPreviewBlock" style="display:none;">
                                        <div style="margin-top:6px; font-size:0.68rem; color:var(--text-secondary);">
                                            <strong>Предпросмотр (первые 5 строк):</strong>
                                        </div>
                                        <div class="import-preview" id="excelPreview"></div>
                                    </div>
                                    <div style="display:flex; gap:6px; margin-top:8px; flex-wrap:wrap;">
                                        <button class="btn btn-warning btn-sm" id="applyExcelImportBtn" disabled>✓ Применить</button>
                                        <button class="btn btn-outline btn-sm" id="cancelExcelImportBtn">✕ Отмена</button>
                                    </div>
                                </div>
                            </div>

                            <div class="io-block">
                                <strong>💾 JSON (резервная копия)</strong>
                                <p>Полная резервная копия расписания и истории.</p>
                                <div style="display:flex; gap:6px; flex-wrap:wrap;">
                                    <button class="btn btn-primary btn-sm" id="exportScheduleBtn">💾 Расписание</button>
                                    <button class="btn btn-outline btn-sm" id="exportHistoryBtn">📜 История</button>
                                    <button class="btn btn-warning btn-sm" id="importBtn">📂 Импорт</button>
                                    <button class="btn btn-outline btn-sm" id="copyJsonBtn">📋 Копировать</button>
                                </div>
                                <input type="file" id="importFile" accept=".json,application/json" style="display:none;">
                                <div id="importFileName" style="margin-top:6px; font-size:0.68rem; color:var(--text-muted);"></div>
                            </div>
                        </div>

                        <div>
                            <div class="io-block">
                                <strong>📘 Документация</strong>
                                <p>Скачать инструкцию в формате Word (.docx).</p>
                                <div style="display:flex; gap:6px; flex-wrap:wrap;">
                                    <button class="btn btn-primary btn-sm" id="downloadAdminGuideBtn">📘 Полная</button>
                                    <button class="btn btn-outline btn-sm" id="downloadShortGuideBtn">📄 Краткая</button>
                                </div>
                                <div id="guideStatus" style="margin-top:6px; font-size:0.68rem; color:var(--text-muted);"></div>
                            </div>

                            <div class="io-block">
                                <strong>💡 Подсказки</strong>
                                <ul style="margin:0; padding-left:16px; font-size:0.68rem; color:var(--text-secondary); line-height:1.5;">
                                    <li>Делайте <strong>экспорт JSON</strong> перед массовыми правками.</li>
                                    <li>Файлы скачиваются в папку загрузок браузера.</li>
                                    <li>Импорт <strong>заменяет</strong> текущее расписание.</li>
                                    <li>Откат — через вкладку <strong>📜 История</strong>.</li>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="console-tab-content" data-tab="history">
                    <div class="console-toolbar">
                        <button class="btn btn-outline btn-sm" id="undoLastConsoleBtn">↶ Отменить последнее</button>
                        <span class="spacer"></span>
                        <button class="btn btn-danger btn-sm" id="clearHistoryBtn">🗑 Очистить историю</button>
                    </div>
                    <div class="history-list" id="consoleHistoryList"></div>
                </div>

                <div class="console-tab-content" data-tab="security">
                    <div id="security-tab-grid">
                        <div>
                            <div class="security-block">
                                <h3>🔐 Смена пароля администратора</h3>
                                <p>Пароль используется для входа в режим редактирования и консоль. По умолчанию — <strong>sever2</strong>.</p>

                                <div class="password-field">
                                    <label for="currentPasswordInput">Текущий пароль</label>
                                    <input type="password" id="currentPasswordInput" placeholder="Текущий пароль" autocomplete="off">
                                    <button class="toggle-visibility" data-target="currentPasswordInput" title="Показать/скрыть">👁</button>
                                </div>

                                <div class="password-field">
                                    <label for="newPasswordInput">Новый пароль</label>
                                    <input type="password" id="newPasswordInput" placeholder="Минимум 4 символа" autocomplete="off">
                                    <button class="toggle-visibility" data-target="newPasswordInput" title="Показать/скрыть">👁</button>
                                    <div class="password-strength" id="passwordStrength">
                                        <div class="bar"></div>
                                    </div>
                                    <div class="password-strength-label" id="passwordStrengthLabel">Введите пароль</div>
                                </div>

                                <div class="password-field">
                                    <label for="confirmPasswordInput">Подтверждение</label>
                                    <input type="password" id="confirmPasswordInput" placeholder="Повторите пароль" autocomplete="off">
                                    <button class="toggle-visibility" data-target="confirmPasswordInput" title="Показать/скрыть">👁</button>
                                </div>

                                <div class="password-msg" id="passwordChangeMsg"></div>

                                <div style="display:flex; gap:6px; flex-wrap:wrap; margin-top:4px;">
                                    <button class="btn btn-primary btn-sm" id="changePasswordBtn">🔐 Сменить</button>
                                    <button class="btn btn-outline btn-sm" id="resetPasswordBtn">↺ Сбросить к «sever2»</button>
                                </div>

                                <div class="password-info">
                                    <strong>💡</strong> Пароль хранится в браузере. При очистке хранилища вернётся к <strong>sever2</strong>.
                                </div>
                            </div>
                        </div>

                        <div>
                            <div class="security-block">
                                <h3>⚠ Забыли пароль?</h3>
                                <p>Сброс к стандартному <strong>sever2</strong> без ввода текущего пароля. Требуется подтверждение.</p>
                                <button class="btn btn-warning btn-sm" id="forceResetPasswordBtn">🔓 Сбросить к «sever2»</button>
                            </div>

                            <div class="security-block">
                                <h3>📋 Информация о текущем пароле</h3>
                                <div id="currentPasswordInfo" style="font-size:0.72rem; color:var(--text-secondary); line-height:1.55;"></div>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="console-tab-content" data-tab="constructor">
                    <div class="constructor-wrap">

                        <div class="constructor-toolbar">
                            <span class="filter-label">🎨 Конструктор сайта</span>
                            <span class="spacer" style="flex:1;"></span>
                            <button class="btn btn-success btn-sm" id="constrSaveBtn">💾 Применить</button>
                            <button class="btn btn-outline btn-sm" id="constrExportBtn">📤 Экспорт</button>
                            <button class="btn btn-outline btn-sm" id="constrImportBtn">📥 Импорт</button>
                            <button class="btn btn-outline btn-sm" id="constrResetBtn">↺ Сброс</button>
                            <button class="btn btn-danger btn-sm" id="constrFactoryBtn">🏭 Заводские</button>
                            <input type="file" id="constrImportFile" accept=".json" style="display:none;">
                        </div>

                        <div class="constructor-layout">
                            <aside class="constructor-nav" id="constructorNav">
                                <button class="constr-nav-btn active" data-section="header">🏷️ Шапка</button>
                                <button class="constr-nav-btn" data-section="tabs">📑 Вкладки</button>
                                <button class="constr-nav-btn" data-section="theme">🎨 Цвета</button>
                                <button class="constr-nav-btn" data-section="sizes">📐 Размеры</button>
                                <button class="constr-nav-btn" data-section="effects">✨ Эффекты</button>
                                <button class="constr-nav-btn" data-section="buttons">🔘 Кнопки</button>
                                <button class="constr-nav-btn" data-section="footer">📌 Футер</button>
                                <button class="constr-nav-btn" data-section="customTabs">🧩 Свои вкладки</button>
                            </aside>

                            <div class="constructor-content" id="constructorContent">

                                <div class="constr-section active" data-section="header">
                                    <h3>🏷️ Шапка сайта</h3>
                                    <div class="constr-field">
                                        <label>Заголовок (H1)</label>
                                        <input type="text" id="constrHeaderTitle" maxlength="120">
                                    </div>
                                    <div class="constr-field">
                                        <label>Подзаголовок (строка 2)</label>
                                        <input type="text" id="constrHeaderSubtitle" maxlength="200">
                                    </div>
                                    <div class="constr-field">
                                        <label>Третья строка (годы/подпись)</label>
                                        <input type="text" id="constrHeaderNote" maxlength="100">
                                    </div>
                                    <div class="constr-field-row">
                                        <label class="constr-check"><input type="checkbox" id="constrShowHeaderTitle"> Показывать заголовок</label>
                                        <label class="constr-check"><input type="checkbox" id="constrShowHeaderSubtitle"> Показывать подзаголовок</label>
                                        <label class="constr-check"><input type="checkbox" id="constrShowHeaderNote"> Показывать третью строку</label>
                                    </div>
                                    <div class="constr-field">
                                        <label>Цвет заголовка</label>
                                        <div class="constr-color-row">
                                            <input type="color" id="constrHeaderTitleColor">
                                            <input type="text" id="constrHeaderTitleColorHex" maxlength="7" placeholder="#0b2a4a">
                                        </div>
                                    </div>
                                </div>

                                <div class="constr-section" data-section="tabs">
                                    <h3>📑 Вкладки расписания</h3>
                                    <p class="constr-hint">Переименовывайте, скрывайте, перемещайте. Основную вкладку «Учительское» удалить нельзя.</p>
                                    <div id="constrTabsList" class="constr-tabs-list"></div>
                                    <div style="margin-top:16px; padding-top:12px; border-top:1px dashed var(--border);">
                                        <h4 style="font-size:0.85rem; margin-bottom:8px;">➕ Добавить новую вкладку</h4>
                                        <div class="constr-field-row">
                                            <div class="constr-field" style="flex:1;">
                                                <label>Заголовок</label>
                                                <input type="text" id="constrNewTabLabel" placeholder="Например: Кружки">
                                            </div>
                                            <div class="constr-field" style="flex:0 0 80px;">
                                                <label>Иконка</label>
                                                <input type="text" id="constrNewTabIcon" placeholder="🎯" maxlength="4">
                                            </div>
                                        </div>
                                        <div class="constr-field">
                                            <label>Содержимое (HTML)</label>
                                            <textarea id="constrNewTabContent" rows="6" placeholder="<h2>Добро пожаловать</h2><p>Ваш текст...</p>"></textarea>
                                        </div>
                                        <button class="btn btn-success btn-sm" id="constrAddTabBtn">➕ Добавить вкладку</button>
                                    </div>
                                </div>

                                <div class="constr-section" data-section="theme">
                                    <h3>🎨 Цвета и темы</h3>
                                    <div class="constr-field">
                                        <label>Быстрые пресеты</label>
                                        <div class="constr-presets">
                                            <button class="constr-preset-btn" data-preset="classic">Классика</button>
                                            <button class="constr-preset-btn" data-preset="dark">Тёмная</button>
                                            <button class="constr-preset-btn" data-preset="sea">Морская</button>
                                            <button class="constr-preset-btn" data-preset="forest">Лесная</button>
                                            <button class="constr-preset-btn" data-preset="sunset">Закат</button>
                                            <button class="constr-preset-btn" data-preset="minimal">Минимализм</button>
                                            <button class="constr-preset-btn" data-preset="purple">Фиолетовая</button>
                                        </div>
                                    </div>
                                    <div class="constr-color-grid">
                                        <div class="constr-field">
                                            <label>Акцент</label>
                                            <div class="constr-color-row">
                                                <input type="color" id="constrColorAccent">
                                                <input type="text" id="constrColorAccentHex" maxlength="7">
                                            </div>
                                        </div>
                                        <div class="constr-field">
                                            <label>Фон страницы</label>
                                            <div class="constr-color-row">
                                                <input type="color" id="constrColorBgPage">
                                                <input type="text" id="constrColorBgPageHex" maxlength="7">
                                            </div>
                                        </div>
                                        <div class="constr-field">
                                            <label>Фон контейнера</label>
                                            <div class="constr-color-row">
                                                <input type="color" id="constrColorBgContainer">
                                                <input type="text" id="constrColorBgContainerHex" maxlength="7">
                                            </div>
                                        </div>
                                        <div class="constr-field">
                                            <label>Основной текст</label>
                                            <div class="constr-color-row">
                                                <input type="color" id="constrColorTextPrimary">
                                                <input type="text" id="constrColorTextPrimaryHex" maxlength="7">
                                            </div>
                                        </div>
                                        <div class="constr-field">
                                            <label>Вторичный текст</label>
                                            <div class="constr-color-row">
                                                <input type="color" id="constrColorTextSecondary">
                                                <input type="text" id="constrColorTextSecondaryHex" maxlength="7">
                                            </div>
                                        </div>
                                        <div class="constr-field">
                                            <label>Границы</label>
                                            <div class="constr-color-row">
                                                <input type="color" id="constrColorBorder">
                                                <input type="text" id="constrColorBorderHex" maxlength="7">
                                            </div>
                                        </div>
                                        <div class="constr-field">
                                            <label>Замена (фон)</label>
                                            <div class="constr-color-row">
                                                <input type="color" id="constrColorSubstBg">
                                                <input type="text" id="constrColorSubstBgHex" maxlength="7">
                                            </div>
                                        </div>
                                        <div class="constr-field">
                                            <label>Замена (текст)</label>
                                            <div class="constr-color-row">
                                                <input type="color" id="constrColorSubstText">
                                                <input type="text" id="constrColorSubstTextHex" maxlength="7">
                                            </div>
                                        </div>
                                        <div class="constr-field">
                                            <label>Успех</label>
                                            <div class="constr-color-row">
                                                <input type="color" id="constrColorSuccess">
                                                <input type="text" id="constrColorSuccessHex" maxlength="7">
                                            </div>
                                        </div>
                                        <div class="constr-field">
                                            <label>Опасность</label>
                                            <div class="constr-color-row">
                                                <input type="color" id="constrColorDanger">
                                                <input type="text" id="constrColorDangerHex" maxlength="7">
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <div class="constr-section" data-section="sizes">
                                    <h3>📐 Размеры и типографика</h3>
                                    <div class="constr-field">
                                        <label>Размер основного текста: <span id="constrSizeBaseVal">0.78</span>rem</label>
                                        <input type="range" id="constrSizeBase" min="0.6" max="1.2" step="0.02">
                                    </div>
                                    <div class="constr-field">
                                        <label>Размер заголовка школы: <span id="constrSizeTitleVal">1.35</span>rem</label>
                                        <input type="range" id="constrSizeTitle" min="0.9" max="2.5" step="0.05">
                                    </div>
                                    <div class="constr-field">
                                        <label>Скругление блоков: <span id="constrRadiusVal">12</span>px</label>
                                        <input type="range" id="constrRadius" min="0" max="30" step="1">
                                    </div>
                                    <div class="constr-field">
                                        <label>Ширина контейнера: <span id="constrContainerWidthVal">1800</span>px</label>
                                        <input type="range" id="constrContainerWidth" min="900" max="2400" step="20">
                                    </div>
                                    <div class="constr-field">
                                        <label>Отступы контейнера: <span id="constrContainerPaddingVal">20</span>px</label>
                                        <input type="range" id="constrContainerPadding" min="6" max="40" step="2">
                                    </div>
                                    <div class="constr-field">
                                        <label>Плотность строк таблицы: <span id="constrRowPaddingVal">6</span>px</label>
                                        <input type="range" id="constrRowPadding" min="2" max="16" step="1">
                                    </div>
                                </div>

                                <div class="constr-section" data-section="effects">
                                    <h3>✨ Эффекты</h3>
                                    <div class="constr-field-row">
                                        <label class="constr-check"><input type="checkbox" id="constrEffectBg"> Анимированный фон</label>
                                        <label class="constr-check"><input type="checkbox" id="constrEffectShadow"> Тень контейнера</label>
                                        <label class="constr-check"><input type="checkbox" id="constrEffectBlur"> Blur (стекло)</label>
                                        <label class="constr-check"><input type="checkbox" id="constrEffectTransitions"> Плавные переходы</label>
                                    </div>
                                    <div class="constr-field">
                                        <label>Интенсивность анимированного фона: <span id="constrBgIntensityVal">1.0</span>×</label>
                                        <input type="range" id="constrBgIntensity" min="0.2" max="2" step="0.1">
                                    </div>
                                </div>

                                <div class="constr-section" data-section="buttons">
                                    <h3>🔘 Кнопки в шапке</h3>
                                    <div class="constr-field-row">
                                        <label class="constr-check"><input type="checkbox" id="constrBtnEdit"> ✎ Редактировать</label>
                                        <label class="constr-check"><input type="checkbox" id="constrBtnMySchedule"> 👤 Моё расписание</label>
                                        <label class="constr-check"><input type="checkbox" id="constrBtnConsole"> ⚙ Консоль</label>
                                        <label class="constr-check"><input type="checkbox" id="constrBtnTheme"> ☀️/🌙 Тема</label>
                                        <label class="constr-check"><input type="checkbox" id="constrBtnMobile"> 📱 Мобильная версия</label>
                                        <label class="constr-check"><input type="checkbox" id="constrBtnIndicators"> 📌 Индикаторы замен и режима</label>
                                        <label class="constr-check"><input type="checkbox" id="constrBtnAdminIndicator"> 👮 Индикатор администратора</label>
                                    </div>
                                </div>

                                <div class="constr-section" data-section="footer">
                                    <h3>📌 Футер</h3>
                                    <div class="constr-field">
                                        <label>Текст подсказки по умолчанию</label>
                                        <input type="text" id="constrFooterHint" maxlength="200">
                                    </div>
                                    <div class="constr-field-row">
                                        <label class="constr-check"><input type="checkbox" id="constrShowFooterStatus"> Показывать бейдж «Сохранено локально»</label>
                                        <label class="constr-check"><input type="checkbox" id="constrShowFooterHistory"> Показывать кнопку «📜 История»</label>
                                    </div>
                                </div>

                                <div class="constr-section" data-section="customTabs">
                                    <h3>🧩 Пользовательские вкладки</h3>
                                    <p class="constr-hint">Вкладки с произвольным HTML-содержимым. Они появятся в шапке рядом с основными.</p>
                                    <div id="constrCustomTabsList"></div>
                                    <div style="margin-top:16px; padding:12px; background:var(--accent-lighter); border-radius:10px; border:1px solid var(--border);">
                                        <h4 style="font-size:0.8rem; margin-bottom:8px;">➕ Новая пользовательская вкладка</h4>
                                        <div class="constr-field-row">
                                            <div class="constr-field" style="flex:1;">
                                                <label>Заголовок</label>
                                                <input type="text" id="constrCustomTabLabel" placeholder="Информация">
                                            </div>
                                            <div class="constr-field" style="flex:0 0 80px;">
                                                <label>Иконка</label>
                                                <input type="text" id="constrCustomTabIcon" placeholder="ℹ️" maxlength="4">
                                            </div>
                                        </div>
                                        <div class="constr-field">
                                            <label>HTML-содержимое</label>
                                            <textarea id="constrCustomTabContent" rows="8" placeholder="<h2>Заголовок</h2><p>Текст с <strong>разметкой</strong>.</p>"></textarea>
                                        </div>
                                        <button class="btn btn-success btn-sm" id="constrAddCustomTabBtn">➕ Добавить</button>
                                    </div>
                                </div>

                            </div>
                        </div>

                        <div class="constr-status" id="constrStatus"></div>
                    </div>
                </div>

                <div class="console-tab-content" data-tab="update">
                    <div class="update-wrap">

                        <div class="update-toolbar">
                            <span class="filter-label">🔄 Управление версией приложения</span>
                            <span class="spacer" style="flex:1;"></span>
                            <button class="btn btn-primary btn-sm" id="updateDownloadHtmlBtn">📤 Скачать HTML</button>
                            <button class="btn btn-outline btn-sm" id="updateCopyHtmlBtn">📋 Копировать HTML</button>
                            <button class="btn btn-outline btn-sm" id="updateBackupDataBtn">💾 Бэкап данных</button>
                        </div>

                        <div class="update-grid">

                            <div class="update-card">
                                <div class="update-card-header">
                                    <span>📥 Загрузить обновление</span>
                                </div>
                                <div class="update-card-body">
                                    <p class="update-hint">
                                        Выберите новый HTML-файл приложения. <strong>Все текущие данные</strong>
                                        (расписания, журнал замен, справочники, настройки) <strong>сохранятся</strong>.
                                    </p>

                                    <div class="update-drop-zone" id="updateDropZone">
                                        <div class="update-drop-icon">📁</div>
                                        <div class="update-drop-text">
                                            <strong>Перетащите HTML-файл</strong> или
                                            <button type="button" class="update-drop-btn" id="updateSelectFileBtn">выберите файл</button>
                                        </div>
                                        <input type="file" id="updateHtmlFile" accept=".html,.htm" style="display:none;">
                                    </div>

                                    <div class="update-file-info" id="updateFileInfo" style="display:none;">
                                        <div class="update-file-row">
                                            <span class="update-file-label">📄 Файл:</span>
                                            <strong id="updateFileName">—</strong>
                                        </div>
                                        <div class="update-file-row">
                                            <span class="update-file-label">📦 Размер:</span>
                                            <span id="updateFileSize">—</span>
                                        </div>
                                        <div class="update-file-row">
                                            <span class="update-file-label">📅 Дата:</span>
                                            <span id="updateFileDate">—</span>
                                        </div>
                                        <div class="update-file-row">
                                            <span class="update-file-label">🏷️ Версия:</span>
                                            <span id="updateFileVersion">не указана</span>
                                        </div>
                                    </div>

                                    <div class="update-checks" id="updateChecks" style="display:none;">
                                        <div class="update-check-item" data-check="doctype">
                                            <span class="update-check-icon">⏳</span>
                                            <span class="update-check-text">Проверка DOCTYPE...</span>
                                        </div>
                                        <div class="update-check-item" data-check="html">
                                            <span class="update-check-icon">⏳</span>
                                            <span class="update-check-text">Проверка тега &lt;html&gt;...</span>
                                        </div>
                                        <div class="update-check-item" data-check="schedule">
                                            <span class="update-check-icon">⏳</span>
                                            <span class="update-check-text">Проверка наличия расписания...</span>
                                        </div>
                                        <div class="update-check-item" data-check="console">
                                            <span class="update-check-icon">⏳</span>
                                            <span class="update-check-text">Проверка консоли...</span>
                                        </div>
                                        <div class="update-check-item" data-check="script">
                                            <span class="update-check-icon">⏳</span>
                                            <span class="update-check-text">Проверка script-блоков...</span>
                                        </div>
                                    </div>

                                    <div class="update-error" id="updateError"></div>

                                    <div class="update-actions" id="updateActions" style="display:none;">
                                        <button class="btn btn-warning btn-sm" id="updateApplyBtn" disabled>✅ Применить обновление</button>
                                        <button class="btn btn-outline btn-sm" id="updateCancelBtn">✕ Отмена</button>
                                    </div>
                                </div>
                            </div>

                            <div class="update-side">
                                <div class="update-card">
                                    <div class="update-card-header">
                                        <span>ℹ️ Информация о версии</span>
                                    </div>
                                    <div class="update-card-body">
                                        <div class="update-info-row">
                                            <span>📅 Размер текущего файла:</span>
                                            <strong id="updateCurrentSize">—</strong>
                                        </div>
                                        <div class="update-info-row">
                                            <span>🏷️ Текущая версия:</span>
                                            <strong id="updateCurrentVersion">—</strong>
                                        </div>
                                        <div class="update-info-row">
                                            <span>🕓 Последнее обновление:</span>
                                            <strong id="updateLastUpdate">не производилось</strong>
                                        </div>
                                        <div class="update-info-row">
                                            <span>💾 Размер данных:</span>
                                            <strong id="updateDataSize">—</strong>
                                        </div>
                                    </div>
                                </div>

                                <div class="update-card">
                                    <div class="update-card-header">
                                        <span>⏪ Откат обновления</span>
                                    </div>
                                    <div class="update-card-body">
                                        <p class="update-hint">
                                            Сохранена резервная копия предыдущей версии HTML.
                                            Можно восстановить, если новое обновление что-то сломало.
                                        </p>
                                        <div class="update-backup-info" id="updateBackupInfo">
                                            <span>Резервная копия:</span> <strong id="updateBackupStatus">не найдена</strong>
                                        </div>
                                        <div style="display:flex; gap:6px; flex-wrap:wrap; margin-top:8px;">
                                            <button class="btn btn-warning btn-sm" id="updateRollbackBtn" disabled>⏪ Откатить</button>
                                            <button class="btn btn-outline btn-sm" id="updateClearBackupBtn" disabled>🗑 Удалить бэкап</button>
                                        </div>
                                    </div>
                                </div>

                                <div class="update-card">
                                    <div class="update-card-header">
                                        <span>🛡️ Автоматический бэкап данных</span>
                                    </div>
                                    <div class="update-card-body">
                                        <p class="update-hint">
                                            Перед обновлением автоматически создаётся JSON-бэкап всех данных.
                                            Если что-то пойдёт не так — можно восстановить.
                                        </p>
                                        <div class="update-backup-info" id="updateDataBackupInfo">
                                            <span>Последний бэкап:</span> <strong id="updateDataBackupStatus">не производился</strong>
                                        </div>
                                        <button class="btn btn-outline btn-sm" id="updateRestoreDataBtn" style="margin-top:8px;" disabled>
                                            📂 Восстановить из бэкапа
                                        </button>
                                    </div>
                                </div>
                            </div>

                        </div>

                        <div class="update-status" id="updateStatus"></div>
                    </div>
                </div>

                <div class="console-tab-content" data-tab="danger">
                    <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px; align-items: start;">
                        <div style="padding:12px; background:#fff5f5; border:1.5px solid var(--danger); border-radius:10px;">
                            <strong style="color:var(--danger); font-size:0.82rem;">⚠ Внимание!</strong>
                            <p style="color:var(--text-secondary); font-size:0.7rem; margin:6px 0 10px; line-height:1.4;">
                                Действия в этом разделе <strong>необратимы</strong>. Рекомендуется сначала сохранить резервную копию JSON.
                            </p>
                            <div style="display:flex; gap:6px; flex-wrap:wrap;">
                                <button class="btn btn-danger btn-sm" id="fullResetBtn">🗑 Полный сброс</button>
                                <button class="btn btn-warning btn-sm" id="clearAllLocalBtn">🧹 Очистить хранилище</button>
                            </div>
                        </div>

                        <div style="padding:12px; background:var(--accent-lighter); border-radius:10px; border:1px solid var(--border);">
                            <strong style="color:var(--text-primary); font-size:0.8rem;">ℹ️ Информация о хранилище</strong>
                            <div id="storageInfo" style="margin-top:6px; font-size:0.72rem; color:var(--text-secondary); line-height:1.6;"></div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div class="modal-overlay" id="teacherEditModal">
        <div class="modal modal-narrow">
            <h2 id="teacherEditTitle">Добавить учителя</h2>
            <p id="teacherEditDesc">Заполните поля и нажмите «Сохранить».</p>

            <label for="teacherNameInput">Имя учителя *</label>
            <input type="text" id="teacherNameInput" placeholder="Иванов И.И.">

            <label for="teacherRoomInput" style="margin-top:8px;">Кабинет</label>
            <input type="text" id="teacherRoomInput" placeholder="101">

            <label for="teacherClassInput" style="margin-top:8px;">Класс</label>
            <input type="text" id="teacherClassInput" placeholder="7а">

            <span class="error-msg" id="teacherEditError"></span>

            <div class="modal-actions">
                <button class="btn btn-outline" id="cancelTeacherEditBtn">Отмена</button>
                <button class="btn btn-primary" id="saveTeacherBtn">Сохранить</button>
            </div>
        </div>
    </div>

    <div class="modal-overlay" id="dictManageModal">
        <div class="modal" style="max-width:640px;">
            <div class="console-header">
                <h2>⚙ Справочники журнала замен</h2>
                <button class="btn btn-outline btn-sm" id="closeDictManageBtn">✕ Закрыть</button>
            </div>

            <div class="console-tabs">
                <button class="console-tab active" data-dict-tab="teachers">👥 Учителя</button>
                <button class="console-tab" data-dict-tab="subjects">📚 Предметы</button>
            </div>

            <div class="dict-body">
                <div class="dict-content active" data-dict-tab="teachers">
                    <div class="dict-hint">Учителя автоматически берутся из «Учительского» расписания. Здесь можно добавить дополнительных (например, совместителей).</div>
                    <div class="dict-add-row">
                        <input type="text" id="dictTeacherInput" placeholder="Иванов И.И." maxlength="80">
                        <button class="btn btn-success btn-sm" id="dictAddTeacherBtn">➕ Добавить</button>
                    </div>
                    <div class="dict-list" id="dictTeacherList"></div>
                </div>

                <div class="dict-content" data-dict-tab="subjects">
                    <div class="dict-hint">Предметы автоматически собраны из «Детского» расписания. Здесь можно добавить свои (например, «Классный час»).</div>
                    <div class="dict-add-row">
                        <input type="text" id="dictSubjectInput" placeholder="Математика" maxlength="80">
                        <button class="btn btn-success btn-sm" id="dictAddSubjectBtn">➕ Добавить</button>
                    </div>
                    <div class="dict-list" id="dictSubjectList"></div>
                </div>
            </div>
        </div>
    </div>

    <div class="modal-overlay" id="myScheduleModal">
        <div class="modal modal-wide">
            <div class="console-header">
                <h2>👤 Моё расписание</h2>
                <button class="btn btn-outline btn-sm" id="closeMyScheduleBtn">✕ Закрыть</button>
            </div>

            <div class="my-sched-controls">
                <div class="my-sched-control">
                    <label for="myScheduleTeacher">Учитель:</label>
                    <select id="myScheduleTeacher" class="teacher-select"></select>
                </div>
                <div class="my-sched-control">
                    <label for="myScheduleDay">День:</label>
                    <select id="myScheduleDay" class="teacher-select">
                        <option value="today">Сегодня</option>
                        <option value="tomorrow">Завтра</option>
                        <option value="ПОНЕДЕЛЬНИК">Понедельник</option>
                        <option value="ВТОРНИК">Вторник</option>
                        <option value="СРЕДА">Среда</option>
                        <option value="ЧЕТВЕРГ">Четверг</option>
                        <option value="ПЯТНИЦА">Пятница</option>
                        <option value="week" selected>Вся неделя</option>
                    </select>
                </div>
                <div class="my-sched-control">
                    <button class="btn btn-primary btn-sm" id="printMyScheduleBtn">🖨 Распечатать</button>
                    <button class="btn btn-outline btn-sm" id="bellsSettingsBtn" title="Расписание звонков">🔔 Звонки</button>
                </div>
            </div>

            <div class="console-body">
                <div id="myScheduleContent"></div>
            </div>
        </div>
    </div>

    <div class="modal-overlay" id="bellsModal">
        <div class="modal modal-narrow">
            <h2>🔔 Расписание звонков</h2>
            <p>Укажите время начала и конца каждого урока.</p>
            <div id="bellsEditor" style="display:grid; grid-template-columns: 60px 1fr 1fr; gap:8px; align-items:center; max-height:300px; overflow-y:auto; margin-bottom:12px;"></div>
            <span class="error-msg" id="bellsError"></span>
            <div class="modal-actions">
                <button class="btn btn-outline" id="cancelBellsBtn">Отмена</button>
                <button class="btn btn-primary" id="saveBellsBtn">Сохранить</button>
            </div>
        </div>
    </div>

    <script>
        // ==================== ИСХОДНЫЕ ДАННЫЕ: УЧИТЕЛЯ ====================
        const rawTeachers = [
            { name: "Чумакова А.Ф.", room: "309", cls: "", pn: ["","","","","","","",""], vt: ["","","5бУО","","6аЛ","7г","",""], sr: ["5бУО","","","4гС","9бС","7б","",""], cht: ["","","4а","6аП","4б","7в","7а",""], pt: ["9бС","4г","4в","","","6аП","",""] },
            { name: "Костенко А.А.", room: "212", cls: "6б", pn: ["вн","6б","6б","6к","6к","","7г",""], vt: ["6б","6к","7г","","","","",""], sr: ["6к","6к","6б","6б","7г","","",""], cht: ["","","7г","6к","6к","6б","",""], pt: ["6б","6б","6б","6к","6к","","",""] },
            { name: "Воробьева Т.И.", room: "306", cls: "8б", pn: ["вн","10в","10в","5в","5в","","10вг1",""], vt: ["8б","8б","5в","5в","","","10в",""], sr: ["8б","5в","5в","","","","",""], cht: ["","5в","10в","8б","","","",""], pt: ["10в","8б","5в","","","","",""] },
            { name: "Гайкова Н.Н.", room: "206", cls: "7а", pn: ["вн","7а","7а","6аЛ","5а","5а","",""], vt: ["","","7а","5а","","","",""], sr: ["5а","5а","7а","6аЛ","","","",""], cht: ["6аЛ","5а","7а","","","","",""], pt: ["6аЛ","7а","5а","5а","","","",""] },
            { name: "Дмитриева Н.Н.", room: "308", cls: "8а", pn: ["вн","8а","8а","7г","8г","5бН","9к",""], vt: ["9к","9к","8г","5бН","8а","5бН","",""], sr: ["","","8г","7г","5бН","9к","",""], cht: ["9к","5бН","8г","8а","","5бН","",""], pt: ["9к","8г","","5бН","5бН","8а","",""] },
            { name: "Ващенко Л.Н.", room: "", cls: "", pn: ["","9бС","9а","6аП","6аП","5бУО","5бУО",""], vt: ["9бС","9а","9а","9а","6аП","6аП","",""], sr: ["9а","9бС","9бС","5бУО","6аП","6аП","",""], cht: ["9а","6аП","6аП","5бУО","5бУО","9бС","",""], pt: ["5бУО","5бУО","5бУО","9бС","9бС","","",""] },
            { name: "Сазонова В.Г.", room: "319", cls: "7в", pn: ["вн","7в","7в","5к","5к","","",""], vt: ["5к","7в","10а","10а","10б","10б","",""], sr: ["7в","5к","5к","","","","",""], cht: ["7в","5к","5к","10а","","10б","10бг2",""], pt: ["5к","10а","10а","7в","10б","10б","",""] },
            { name: "Панская Е.С.", room: "208", cls: "8в", pn: ["вн","5б","11а","8в","","11б","11б","11б"], vt: ["8в","5б","5б","11сг12","11а","11а","",""], sr: ["","","","8в","5б","11г12","",""], cht: ["8в","","5б","5б","11б","11б","11б",""], pt: ["8в","11а","11а","5б","5б","","",""] },
            { name: "Лактионова Л.Н.", room: "", cls: "", pn: ["","","6аЛ","9в","9в","","",""], vt: ["9в","","","","","","",""], sr: ["","6аЛ","9в","","","","",""], cht: ["","6аЛ","9в","","9в","","",""], pt: ["","6аЛ","","","","","",""] },
            { name: "Фоменко Н.И.", room: "", cls: "", pn: ["","6аЗПР","6аЗПР","","","","",""], vt: ["6аЗПР","6аЗПР","","","","","",""], sr: ["","","","","6аЗПР","6аЗПР","",""], cht: ["6аЗПР","6аЗПР","","","","","",""], pt: ["6аЗПР","","","","","","",""] },
            { name: "Литвинова О.А.", room: "", cls: "9б", pn: ["вн","9б","9б","6а","6а","7б","",""], vt: ["9б","6а","6а","","","7б","",""], sr: ["6а","6а","7б","7б","9б","","",""], cht: ["","9б","6а","7б","","","",""], pt: ["","7б","6а","6а","","9б","",""] },
            { name: "Кошелева И.В.", room: "", cls: "10б", pn: ["вн","4б","3б","4в","3бВ","","6аЗПР",""], vt: ["10г2","10г2","3а","6б","","6а","",""], sr: ["1б","10г2","4б","6аЗПР","10г2","3бВ","",""], cht: ["3а","6а","10г2","4в","6б","3б","",""], pt: ["","10г2","6а","6б","6аЗПР","","",""] },
            { name: "Реутова Е.Г.", room: "320", cls: "", pn: ["вн","","","","","","",""], vt: ["1а","2в","2а","5к","5а","5в","",""], sr: ["","","","","","","",""], cht: ["5а","5в","2а","2в","5к","5к","",""], pt: ["","","","","","","",""] },
            { name: "Колкова Ю.С.", room: "", cls: "7г", pn: ["вн","7г","3в","9а","9б","9в","",""], vt: ["2в","8г","2а","2б","1а","","",""], sr: ["9бС","8г","7г","9в","9а","9б","",""], cht: ["2б","2а","","2в","3в","","",""], pt: ["8г","7г","9в","9б","","9а","",""] },
            { name: "Чупакова Т.В.", room: "218", cls: "10в", pn: ["вн","5бН","9к","8б","8в","10хс","","10а"], vt: ["","10вг2","8в","8б","","","",""], sr: ["9к","1в","5бН","10бв","10а","","",""], cht: ["8б","8в","10вг1","","","","",""], pt: ["5бН","10вг1","10бв","9к","10а","","",""] },
            { name: "Ермакова А.А.", room: "", cls: "", pn: ["","3б","4г","9а","9б","5б","7а",""], vt: ["","4а","3а","","7б","7в","",""], sr: ["","","","7а","9а","9б","",""], cht: ["5б","4г","3а","7в","5б","3б","7б",""], pt: ["4а","7в","7а","9б","7б","9а","",""] },
            { name: "Калашникова М.В.", room: "", cls: "9в", pn: ["вн","4б","","8б","8в","9в","8а",""], vt: ["11а","11б","8в","5б","","8б","",""], sr: ["","11б","1б","4б","9в","11а",".",""], cht: ["5б","8а","","","5б","","",""], pt: ["11а","","9в","11б","8в","8б","8а",""] },
            { name: "Жукова А.Е.", room: "303", cls: "", pn: ["","4г","4гС","","6к","8а","",""], vt: ["8г","11г2","11г2","6б","","6а","",""], sr: ["8а","8г","11г2","6к","","","",""], cht: ["8а","6а","4г","11г2","6б","6к","",""], pt: ["8г","11г2","6а","6б","4гС","","",""] },
            { name: "Азарова Т.С.", room: "219", cls: "", pn: ["","6аП","5бН","7а","6б","7в","",""], vt: ["6аП","5бН","7в","7а","7а","6б","7в",""], sr: ["7а","7в","5бН","6аП","6б","","",""], cht: ["6б","","7в","7а","5бН","","",""], pt: ["5бН","","7в","7а","6аП","6б","",""] },
            { name: "Польская Т.И.", room: "214", cls: "", pn: ["","8в","8в","5бУО","8а","9б","",""], vt: ["8а","8а","5бУО","","9б","9б","",""], sr: ["9б","8в","8в","8а","8а","","",""], cht: ["","5бУО","9б","8а","8в","8в","",""], pt: ["8а","9б","8в","5бУО","","","",""] },
            { name: "Виноходов Н.Ю.", room: "", cls: "", pn: ["","","","","7в","7а","7б","9б"], vt: ["9бС","9б","","","","","",""], sr: ["","","","9б","7в","7а","",""], cht: ["9б","7б","","","","","",""], pt: ["","","","","","","",""] },
            { name: "Климова А.А.", room: "", cls: "11б", pn: ["вн","9а","11б","10ав","10ав","","9в",""], vt: ["10ав","10ав","9в","9в","","11б","",""], sr: ["11б","9а","10ав","10ав","","9в","",""], cht: ["11б","10ав","10ав","9а","9а","","10вс",""], pt: ["9а","9а","","","9в","11б","9в",""] },
            { name: "Печесклеева Ю.Ю.", room: "312", cls: "", pn: ["вн","5к","5б","9к","9к","8б","11а","11а"], vt: ["5б","11а","5к","8б","9к","9к","",""], sr: ["5к","11а","8б","8б","9к","5б","8б",""], cht: ["5к","8б","9к","11а","11а","11а","",""], pt: ["5б","5б","8б","11а","11а","5к","",""] },
            { name: "Олейникова О.А.", room: "", cls: "", pn: ["","6аЛ","5в","","","9бС","5а",""], vt: ["5в","5а","5вВ","9бС","","","",""], sr: ["","","6аЛ","5в","5а","5а","",""], cht: ["","","","5в","5а","6аЛ","",""], pt: ["5а","5вВ","9бС","6аЛ","5в","5в","",""] },
            { name: "Ястребова Ю.В.", room: "", cls: "6к", pn: ["вн","6к","7г","10б","10б","8г","8г","8г"], vt: ["7б","7б","","7г","6к","8г","",""], sr: ["7б","7б","10б","10б","8г","6к","",""], cht: ["7б","10б","6к","8г","7г","7г","",""], pt: ["7б","6к","","7г","7г","8г","",""] },
            { name: "Курко О.А.", room: "206", cls: "9к", pn: ["вн","9к","9в","11ф","11ф","7г","9а",""], vt: ["11б","9в","11ф","10б","10а","8а","9к",""], sr: ["10в","","9а","11ф","11ф","7г","",""], cht: ["10б","9к","9а","9в","10а","10в","",""], pt: ["9в","9к","9а","8а","11б","","",""] },
            { name: "Павлюченко Е.Н.", room: "217", cls: "", pn: ["","10и","10и","11и","11и","","",""], vt: ["","","","","","","",""], sr: ["10и","10и","11и","11и","11и","9а","10и",""], cht: ["","","","9к","","","",""], pt: ["","","","","","","",""] },
            { name: "Шоренков А. И.", room: "", cls: "", pn: ["","5в","5к","5а","","5б","",""], vt: ["2в","2в","2б","5б","4а","","9а",""], sr: ["8а","","4б","4б","7б","7в","7г",""], cht: ["6а","6б","8б","8в","9б","9б","",""], pt: ["7а","9в","9к","8г","8в","8б","8а",""] },
            { name: "Шматова Т.И.", room: "205", cls: "", pn: ["","6а","8б","10вг1","10вг1","11си","8в",""], vt: ["","","11си","","6а","6аЗПР","6аЗПР",""], sr: ["6аЗПР","8б","10вг1","10вг1","8в","8г","6а",""], cht: ["8г","10вг1","6аЗПР","6а","6а","6аЗПР","",""], pt: ["","","","","","","",""] },
            { name: "Бурлуцкая Н.А.", room: "215", cls: "", pn: ["","10и","10и","11х","7а","11ф","10г2","10б"], vt: ["8г","8г","8в","8в","8б","8б","9а",""], sr: ["10и","10и","11фс","9в","9в","9а","10и",""], cht: ["8а","8а","11сг12","10х","9б","9б","10г1",""], pt: ["11б","9в","10вс","7б","10в","7в","7г",""] },
            { name: "Савкина Е.И.", room: "315", cls: "5а", pn: ["вн","","6к","","5бН","6б","5б",""], vt: ["6к","7г","5а","6аП","5бН","5б","11а",""], sr: ["5бН","6б","5а","5к","6к","","",""], cht: ["7г","5б","5а","5к","6аП","5б","",""], pt: ["","5а","5к","5к","6б","7г","11а",""] },
            { name: "Юрченко И.В.", room: "211", cls: "9а", pn: ["вн","10бг2","10бг2","8б","9а","9а","10ах","9к"], vt: ["10г1","10г1","10вс","10в","10в","10а","",""], sr: ["9в","9в","9б","9к","10г1","10в","",""], cht: ["","9а","10г1","10бг2","10бвс","9к","10ах",""], pt: ["10бв","10г1","9б","9в","10а","10в","",""] },
            { name: "Духанина Н.А.", room: "302", cls: "6а", pn: ["вн","8г","6а","7в","7б","8а","8б","7а"], vt: ["7в","6аЛ","8б","6аЗПР","8г","8в","",""], sr: ["8в","7а","6аЗПР","","","8а","",""], cht: ["8б","7в","7б","6аЗПР","6аЛ","6а","",""], pt: ["","6а","8г","8в","8а","7а","7б",""] },
            { name: "Бавтрукович В.О.", room: "", cls: "5в", pn: ["вн","11х","","11сг12","11сг12","","",""], vt: ["","5в","11г1","11ифх","9бС","","",""], sr: ["5в","11г2","11г1","11сг12","11сг12","11х","",""], cht: ["5в","9бС","11ифх","11г1","","","",""], pt: ["5в","11г1","","11г2","","","",""] },
            { name: "Андреева А.А.", room: "318", cls: "10а", pn: ["вн","","5вВ","7б","7г","10а","7в",""], vt: ["","","","5вВ","11б","5вВ","7а",""], sr: ["","","5б","5а","5в","5к","10в",""], cht: ["5бН","11а","5вВ","","7а","","8а",""], pt: ["7г","8а","10б","","","7б","7в",""] },
            { name: "Жиренко Н.В.", room: "", cls: "", pn: ["","8б","6аП","8г","9бС","8в","9б",""], vt: ["","","6аЛ","9к","6аЗПР","9а","9в",""], sr: ["8г","","6аП","9бС","","","9б",""], cht: ["9в","8в","6аЛ","","9к","8б","9а",""], pt: ["","","","","","","",""] },
            { name: "Попова Т.Г.", room: "209", cls: "8г", pn: ["вн","11а","8г","6б","11х","5к","6к",""], vt: ["","6аП","8а","6аЛ","5в","5а","",""], sr: ["","","11х","11х","8б","8в","",""], cht: ["7а","7г","8а","6аЛ","7б","8г","",""], pt: ["7в","11х","11бг12","6аЗПР","8б","8в","",""] },
            { name: "Бабукова В.В.", room: "", cls: "", pn: ["","10х","9бС","","","","6а","10в"], vt: ["9а","10х","9к","9б","5бУО","9в","10б",""], sr: ["5б","10х","9к","","","","",""], cht: ["10а","9в","9б","","","9а","",""], pt: ["10х","10ис","5бН","10бв","5бУО","9бС","",""] },
            { name: "Шинкаренко Л.Я.", room: "310", cls: "", pn: ["","11бг12","10х","8а","8б","10г12","10вс",""], vt: ["10х","8в","11х","8г","9в","","",""], sr: ["9к","9б","8а","9а","11х","8б","",""], cht: ["","8г","10х","11х","8в","9в","11а",""], pt: ["10а","10х","11х","9к","9а","9б","",""] },
            { name: "Ходыкина Л.М.", room: "112", cls: "", pn: ["","5а","5а","5б","5б","5в","5в","8в"], vt: ["7а","7а","7б","7б","5к","5к","",""], sr: ["7г","7г","7в","7в","6а","6а","",""], cht: ["6к","6к","6б","6б","8г","8а","8б",""], pt: ["","","9а","9б","9в","9к","",""] },
            { name: "Добрыденко А.С.", room: "113", cls: "11а", pn: ["вн","","","6аЛ","6аЛ","9бС","9бС",""], vt: ["7а","7а","7б","7б","5к","5к","9бС","6аЛ"], sr: ["7г","7г","7в","7в","","5вВ","",""], cht: ["5в","5б","5а","9бС","9бС","8а","8б",""], pt: ["7а","9бС","5к","8г","6аЛ","6аЛ","6аЛ",""] },
            { name: "вакансия", room: "", cls: "", pn: ["","5а","5а","5б","5б","5в","5в","8в"], vt: ["","","","","","","",""], sr: ["","","","","6а","6а","",""], cht: ["6к","6к","6б","6б","8г","","",""], pt: ["","","","9а","9б","9в","9к",""] },
            { name: "Коновалова А.А.", room: "", cls: "", pn: ["","","","","","","",""], vt: ["5бУО","5вВ","6аЗПР","4в","","7г","7б",""], sr: ["6б","5бУО","2а","6а","4а","5бН","7в",""], cht: ["","5вВ","4б","3а","5в","7а","",""], pt: ["5к","","6к","","5а","5б","",""] },
            { name: "Волкова А.В.", room: "", cls: "7б", pn: ["вн","7б","2а","5вВ","6аЗПР","3а","5бН","8б"], vt: ["7г","5бУО","","","7в","","8г",""], sr: ["","","5вВ","5б","4в","","8а",""], cht: ["","7а","","","","","5а","6б"], pt: ["6а","5в","4а","4б","5к","6к","8в",""] },
            { name: "Лазаренко К.А.", room: "", cls: "", pn: ["","2авС","1аЛ","9бС","3б","6а","5к","8а"], vt: ["3в","5к","6аП","1аЛ","8в","9бС","8б",""], sr: ["","8а","6а","3в","5к","2авС","8в",""], cht: ["2авС","3б","9бС","8б","1аЛ","6аП","",""], pt: ["8б","8в","8а","3б","3в","6а","",""] },
            { name: "Ермоленко А.П.", room: "", cls: "", pn: ["","2б","1а","1б","4г","2в","3вП",""], vt: ["5а","4г","1б","6к","5б","4гЮИДД","3Вп",""], sr: ["2в","5б","6к","8г","2б","5в","3вП",""], cht: ["1б","2в","1а","5а","2б","5в","8г",""], pt: ["6к","1а","5б","5в","8г","5а","",""] },
            { name: "Ряполов Д.Н.", room: "", cls: "", pn: ["","9в","7б","6аЗПР","6аУО","9к","6б","9а"], vt: ["","9б","6б","7в","7г","7а","11б",""], sr: ["11а","6аЗПР","5бУО","","7а","6б","11б",""], cht: ["11а","","","7г","7в","7б","9в",""], pt: ["9б","7б","7г","7в","7а","9к","9а",""] },
            { name: "Кизилов Д.А.", room: "", cls: "", pn: ["","1аЗин","3вЧЯр","1аМер","5вВ","1бВ","6аЛ","1бН"], vt: ["4гД","3аДД","1аМак","1бВ","6аЛ","1бН","4аК",""], sr: ["1аЗин","3вЧЯр","1аМер","1аМак","5вВ","6аЛ","4аК",""], cht: ["","3аДД","4гД","1бВ","3вЧЯр","","",""], pt: ["4гД","1аЗин","1аМак","4аК","5вВ","1аМер","",""] },
            { name: "ФК 1-4 кл", room: "", cls: "", pn: ["","","","4в","","","3а",""], vt: ["2а","","1в","4а 4б","","","",""], sr: ["","2а","","1в","","","",""], cht: ["4а 4б","","4в","","1в 3а","","",""], pt: ["","","3а","2а","","","",""] },
            { name: "Павлюченко А.В.", room: "", cls: "", pn: ["","","","","","","",""], vt: ["6а","6б","6к","6а","6б","6к","6к",""], sr: ["","","","","","","",""], cht: ["6а","6б","","","","","",""], pt: ["","","","","","","",""] },
            { name: "Корнев Н.С.", room: "", cls: "", pn: ["","","","","","","",""], vt: ["9в","","","8а","9а","10в","10а",""], sr: ["10б","10в","","","10а","10б","9к",""], cht: ["10в","11б","8в","","","10а","9б",""], pt: ["","","","8б","10б","11а","8г",""] },
            { name: "Фоменко И.Н.", room: "", cls: "", pn: ["","5бУО","5бУО","","","6аП","6аП",""], vt: ["","","","","","","",""], sr: ["6аП","6аП","","","5аУО","5аУО","",""], cht: ["5аУО","5аУО","5бН","5бН","6аЗПР","","",""], pt: ["6аП","6аП","6аЗПР","","","","",""] },
            { name: "Кориш О.Е.", room: "", cls: "", pn: ["","","","","","5вВ","","5бУО"], vt: ["5вВ","","","","","","","5бН"], sr: ["","","","","","","",""], cht: ["","","","","","","5бУО",""], pt: ["","","","","","","",""] },
            { name: "Кийкова О.В.", room: "", cls: "", pn: ["","","","","","","",""], vt: ["","","","","","","",""], sr: ["","","","","","","9бС",""], cht: ["","","","","","","",""], pt: ["","","","","","","",""] },
            { name: "Хасанова А.В.", room: "", cls: "", pn: ["","","","","","6аЗПР","",""], vt: ["","","","","","","",""], sr: ["","","","","","","",""], cht: ["6аП","","","","","","",""], pt: ["","","","","","","",""] },
            { name: "Куприна А.С.", room: "", cls: "", pn: ["","","","","5бН","","",""], vt: ["5бН","","","5бН","","","5бУО","6аП"], sr: ["","5бН","","","","","6аЗПР",""], cht: ["9бС","","","","","5бУО","",""], pt: ["","","6аП","","","6аЗПР","9бС",""] },
            { name: "Подгорнева А.А.", room: "", cls: "", pn: ["","","","","","","",""], vt: ["","","","","","","",""], sr: ["","","","","","","",""], cht: ["","","","","","","",""], pt: ["5вВ","","","","","","",""] },
            { name: "Гнездилова О.Н.", room: "", cls: "", pn: ["","","","","","","",""], vt: ["","","","","","","",""], sr: ["","5вВ","","","","","",""], cht: ["","","","","","","",""], pt: ["","","","","","","",""] },
            { name: "Мотлохова О.В.", room: "", cls: "", pn: ["","","","","","","","6аП"], vt: ["","9бС","","","","5бУО","",""], sr: ["","","","","9бС","","",""], cht: ["","","","","","6аЗПР","",""], pt: ["6аЗПР","","6аП","","5бУО","","",""] },
            { name: "Гармашова О.А.", room: "", cls: "", pn: ["вн","5вВ","","","","","",""], vt: ["","","","","","","",""], sr: ["5вВ","","","","","","",""], cht: ["","","","5вВ","5вВ","","",""], pt: ["","5вВ","5вВ","","","","",""] },
            { name: "Великих С.В.", room: "", cls: "", pn: ["","","","","","","",""], vt: ["","","","","","","",""], sr: ["","","","","","","",""], cht: ["","","","","","","",""], pt: ["","","","","","","",""] },
            { name: "Назаренко А.Э", room: "", cls: "", pn: ["","","","","","","",""], vt: ["","","","","","","",""], sr: ["","","","","","","",""], cht: ["","","","","","","",""], pt: ["","","","","","","",""] },
            { name: "Аркатова А.А.", room: "", cls: "", pn: ["","","","","","","",""], vt: ["","","","","","","",""], sr: ["","","","","","","",""], cht: ["","","","","","","",""], pt: ["","","","","","","",""] },
        ];

        const rawClasses = [
            { name: "5а", room: "206", lessons: {
                "ПОНЕДЕЛЬНИК": ["Математика", "Русский язык", "Литература", "История", "Английский язык", "Биология", "ИЗО", ""],
                "ВТОРНИК":     ["Математика", "Литература", "Русский язык", "Алгоритмика / английский", "Физкультура", "История", "", ""],
                "СРЕДА":       ["Русский язык", "Литература", "История", "География", "Математика", "Занимательная математика", "", ""],
                "ЧЕТВЕРГ":     ["Английский язык", "Русский язык", "ДНКР / искусственный интеллект", "Физкультура", "Математика", "Музыка", "", ""],
                "ПЯТНИЦА":     ["Математика", "История", "Русский язык", "Литература", "ИЗО", "Физкультура", "", ""]
            }},
            { name: "5б", room: "212", lessons: {
                "ПОНЕДЕЛЬНИК": ["Русский язык", "Математика", "Технология", "История", "Литература", "История", "", ""],
                "ВТОРНИК":     ["Математика", "Литература", "Русский язык", "Алгоритмика / английский", "Физкультура", "История", "", ""],
                "СРЕДА":       ["Биология", "Физкультура", "География", "Музыка", "Русский язык", "Математика", "", ""],
                "ЧЕТВЕРГ":     ["Английский язык на каждый день", "ДНКР / искусственный интеллект", "Литература", "Русский язык", "Английский язык", "История", "", ""],
                "ПЯТНИЦА":     ["Математика", "Математика", "Физкультура", "Русский язык", "Литература", "ИЗО", "", ""]
            }},
            { name: "5в", room: "306", lessons: {
                "ПОНЕДЕЛЬНИК": ["Алгоритмика", "Математика", "Русский язык", "Литература", "Технология", "Технология", "", ""],
                "ВТОРНИК":     ["Математика", "История", "Литература", "Русский язык", "Биология", "Английский язык", "", ""],
                "СРЕДА":       ["История", "Русский язык", "Литература", "Математика", "География", "Физкультура", "", ""],
                "ЧЕТВЕРГ":     ["ДНКР / искусственный интеллект", "Английский язык", "Русский язык", "Математика", "ИЗО", "Физкультура", "", ""],
                "ПЯТНИЦА":     ["История", "Музыка", "Русский язык", "Физкультура", "Математика", "Занимательная математика", "", ""]
            }},
            { name: "5к", room: "319", lessons: {
                "ПОНЕДЕЛЬНИК": ["Математика", "Алгоритмика", "Русский язык", "Литература", "Биология", "Физкультура", "", ""],
                "ВТОРНИК":     ["Русский язык", "Физкультура", "Математика", "Английский язык", "Технология", "Технология", "", ""],
                "СРЕДА":       ["Математика", "Русский язык", "Литература", "История", "Физкультура", "География", "", ""],
                "ЧЕТВЕРГ":     ["Математика", "Русский язык", "Литература", "История", "Английский язык", "Английский на каждый день", "", ""],
                "ПЯТНИЦА":     ["ИЗО", "Русский язык", "ДНКР / искусственный интеллект", "История", "Музыка", "Математика", "", ""]
            }},
            { name: "6а", room: "206", lessons: {
                "ПОНЕДЕЛЬНИК": ["Математика", "История", "Русский язык", "Литература", "Физкультура", "Биология", "", ""],
                "ВТОРНИК":     ["География", "Русский язык", "Русский язык", "ОБЗР", "Математика", "Английский язык", "", ""],
                "СРЕДА":       ["Русский язык", "Литература", "Физкультура", "ИЗО", "Технология", "Технология", "Математика", ""],
                "ЧЕТВЕРГ":     ["Алгоритмика", "Английский язык", "Русский язык", "Математика", "Математика", "История", "", ""],
                "ПЯТНИЦА":     ["Музыка", "История", "Английский язык", "Литература", "Русский язык", "Физкультура", "", ""]
            }},
            { name: "6б", room: "219", lessons: {
                "ПОНЕДЕЛЬНИК": ["Русский язык", "Математика", "Литература", "Биология", "Математика", "Физкультура", "", ""],
                "ВТОРНИК":     ["Русский язык", "География", "Физкультура", "Английский язык", "ОБЗР", "Математика", "", ""],
                "СРЕДА":       ["ИЗО", "История", "Литература", "Русский язык", "Математика", "Физкультура", "", ""],
                "ЧЕТВЕРГ":     ["Математика", "Алгоритмика", "Технология", "Технология", "Русский язык", "Английский язык", "Музыка", ""],
                "ПЯТНИЦА":     ["Русский язык", "Русский язык", "Литература", "Английский язык", "История", "Математика", "", ""]
            }},
            { name: "6к", room: "206", lessons: {
                "ПОНЕДЕЛЬНИК": ["Технология", "Технология", "Математика", "Литература", "Английский язык", "Биология", "Алгоритмика", ""],
                "ВТОРНИК":     ["История", "Русский язык", "География", "Физическая культура", "Математика", "ОБЗР", "", ""],
                "СРЕДА":       ["Русский язык", "Русский язык", "Физическая культура", "Английский язык", "История", "Математика", "", ""],
                "ЧЕТВЕРГ":     ["Технология", "Технология", "Математика", "Русский язык", "Литература", "Английский язык", "", ""],
                "ПЯТНИЦА":     ["Физкультура", "Математика", "ИЗО", "Русский язык", "Литература", "Музыка", "", ""]
            }},
            { name: "7а", room: "206", lessons: {
                "ПОНЕДЕЛЬНИК": ["Русский язык", "Литература", "Алгебра", "Геометрия", "Физика", "История", "Английский язык", ""],
                "ВТОРНИК":     ["Алгебра", "Русский язык", "География", "Физика", "Биология", "Физкультура", "", ""],
                "СРЕДА":       ["Литература", "Русский язык", "История", "Алгебра", "Геометрия", "Английский язык", "", ""],
                "ЧЕТВЕРГ":     ["Алгебра", "Русский язык", "Геометрия", "История", "Физика", "Физкультура", "Английский язык", ""],
                "ПЯТНИЦА":     ["Геометрия", "История", "Литература", "Английский язык", "Биология", "География", "", ""]
            }},
            { name: "7б", room: "219", lessons: {
                "ПОНЕДЕЛЬНИК": ["Алгебра", "Геометрия", "Русский язык", "Литература", "Физика", "История", "Английский язык", ""],
                "ВТОРНИК":     ["География", "Алгебра", "Английский язык", "Биология", "Русский язык", "Литература", "", ""],
                "СРЕДА":       ["Физика", "История", "Русский язык", "Литература", "Алгебра", "Геометрия", "", ""],
                "ЧЕТВЕРГ":     ["Алгебра", "Физика", "Русский язык", "Английский язык", "География", "История", "", ""],
                "ПЯТНИЦА":     ["Литература", "Биология", "Алгебра", "Физика", "Английский язык", "История", "", ""]
            }},
            { name: "7в", room: "319", lessons: {
                "ПОНЕДЕЛЬНИК": ["Русский язык", "Алгебра", "Геометрия", "Физика", "Литература", "История", "Английский язык", ""],
                "ВТОРНИК":     ["Алгебра", "Английский язык", "История", "География", "Биология", "Русский язык", "", ""],
                "СРЕДА":       ["Литература", "Физика", "Русский язык", "Алгебра", "Геометрия", "История", "", ""],
                "ЧЕТВЕРГ":     ["Геометрия", "Алгебра", "Русский язык", "Литература", "Английский язык", "Физика", "", ""],
                "ПЯТНИЦА":     ["История", "Биология", "Литература", "Английский язык", "Алгебра", "Физкультура", "", ""]
            }},
            { name: "7г", room: "303", lessons: {
                "ПОНЕДЕЛЬНИК": ["Алгебра", "Геометрия", "Русский язык", "Литература", "История", "Физика", "Английский язык", ""],
                "ВТОРНИК":     ["История", "Русский язык", "Алгебра", "Биология", "География", "Литература", "", ""],
                "СРЕДА":       ["Физика", "Английский язык", "Русский язык", "Геометрия", "Алгебра", "Литература", "", ""],
                "ЧЕТВЕРГ":     ["Алгебра", "Физика", "История", "География", "Английский язык", "Русский язык", "", ""],
                "ПЯТНИЦА":     ["Литература", "История", "Алгебра", "Физика", "Физкультура", "Биология", "", ""]
            }},
            { name: "8а", room: "308", lessons: {
                "ПОНЕДЕЛЬНИК": ["Алгебра", "Геометрия", "История", "Русский язык", "Литература", "Химия", "Английский язык", ""],
                "ВТОРНИК":     ["Русский язык", "Литература", "Биология", "Алгебра", "Физика", "История", "ОБЗР", ""],
                "СРЕДА":       ["Русский язык", "Физика", "Алгебра", "Геометрия", "Английский язык", "История", "", ""],
                "ЧЕТВЕРГ":     ["История", "Математика", "Русский язык", "Литература", "Алгебра", "География", "Английский язык", ""],
                "ПЯТНИЦА":     ["Алгебра", "Русский язык", "Информатика", "Литература", "ОБЗР", "География", "", ""]
            }},
            { name: "8б", room: "306", lessons: {
                "ПОНЕДЕЛЬНИК": ["Русский язык", "Литература", "Алгебра", "Химия", "Биология", "История", "Английский язык", ""],
                "ВТОРНИК":     ["Алгебра", "Геометрия", "Физика", "Русский язык", "История", "Английский язык", "", ""],
                "СРЕДА":       ["Химия", "Русский язык", "Физика", "Алгебра", "Биология", "Литература", "", ""],
                "ЧЕТВЕРГ":     ["История", "География", "Английский язык", "Литература", "ОБЗР", "Алгебра", "Физкультура", ""],
                "ПЯТНИЦА":     ["География", "Русский язык", "Физика", "История", "Литература", "Математика", "", ""]
            }},
            { name: "8в", room: "208", lessons: {
                "ПОНЕДЕЛЬНИК": ["Русский язык", "Литература", "География", "Химия", "История", "Алгебра", "Английский язык", ""],
                "ВТОРНИК":     ["Английский язык", "Физика", "Алгебра", "История", "Русский язык", "Литература", "", ""],
                "СРЕДА":       ["Русский язык", "История", "Алгебра", "Биология", "География", "Литература", "", ""],
                "ЧЕТВЕРГ":     ["Английский язык", "Биология", "Литература", "Алгебра", "Физика", "География", "", ""],
                "ПЯТНИЦА":     ["Химия", "Русский язык", "История", "Математика", "ОБЗР", "Английский язык", "", ""]
            }},
            { name: "8г", room: "209", lessons: {
                "ПОНЕДЕЛЬНИК": ["Литература", "Русский язык", "Биология", "Физика", "История", "Химия", "Английский язык", ""],
                "ВТОРНИК":     ["Русский язык", "История", "Алгебра", "Химия", "География", "Литература", "", ""],
                "СРЕДА":       ["Английский язык", "Физика", "История", "Химия", "Русский язык", "Биология", "", ""],
                "ЧЕТВЕРГ":     ["История", "Математика", "Физика", "Биология", "География", "Русский язык", "", ""],
                "ПЯТНИЦА":     ["Физика", "Химия", "Русский язык", "Математика", "Английский язык", "История", "", ""]
            }},
            { name: "9а", room: "211", lessons: {
                "ПОНЕДЕЛЬНИК": ["Химия", "География", "История", "Алгебра", "Русский язык", "Литература", "Английский язык", ""],
                "ВТОРНИК":     ["Алгебра", "География", "Геометрия", "Русский язык", "Литература", "История", "", ""],
                "СРЕДА":       ["Русский язык", "Физика", "География", "История", "Геометрия", "Английский язык", "", ""],
                "ЧЕТВЕРГ":     ["Индивидуальный проект", "Математика", "Математика", "Литература", "Физика", "Обществознание", "", ""],
                "ПЯТНИЦА":     ["Химия", "Биология", "Русский язык", "Информатика", "Литература", "История", "", ""]
            }},
            { name: "9б", room: "205", lessons: {
                "ПОНЕДЕЛЬНИК": ["Русский язык", "Математика", "История", "Литература", "Английский язык", "Биология", "Обществознание", ""],
                "ВТОРНИК":     ["Русский язык", "География", "Математика", "Информатика", "История", "Химия", "", ""],
                "СРЕДА":       ["Математика", "Физика", "Русский язык", "Биология", "История", "География", "", ""],
                "ЧЕТВЕРГ":     ["Математика", "Русский язык", "История", "Литература", "Обществознание", "Английский язык", "", ""],
                "ПЯТНИЦА":     ["Физика", "Информатика", "География", "Математика", "Русский язык", "ОБЗР", "", ""]
            }},
            { name: "9в", room: "214", lessons: {
                "ПОНЕДЕЛЬНИК": ["Алгебра", "История", "Русский язык", "Литература", "Английский язык", "Физика", "Химия", ""],
                "ВТОРНИК":     ["Русский язык", "Алгебра", "Информатика", "История", "Биология", "География", "", ""],
                "СРЕДА":       ["Физика", "Химия", "Математика", "Русский язык", "Обществознание", "Английский язык", "", ""],
                "ЧЕТВЕРГ":     ["Русский язык", "Литература", "История", "Математика", "ОБЗР", "Физика", "", ""],
                "ПЯТНИЦА":     ["Химия", "Биология", "Математика", "Русский язык", "Информатика", "География", "", ""]
            }},
            { name: "9к", room: "206", lessons: {
                "ПОНЕДЕЛЬНИК": ["Русский язык", "Математика", "История", "Литература", "Английский язык", "Физика", "Химия", ""],
                "ВТОРНИК":     ["Русский язык", "География", "Математика", "Биология", "История", "Информатика", "", ""],
                "СРЕДА":       ["Математика", "Обществознание", "Русский язык", "Литература", "Физика", "Химия", "", ""],
                "ЧЕТВЕРГ":     ["История", "Математика", "Русский язык", "ОБЗР", "Биология", "География", "", ""],
                "ПЯТНИЦА":     ["Физика", "Химия", "Русский язык", "Математика", "Информатика", "История", "", ""]
            }},
            { name: "10а", room: "318", lessons: {
                "ПОНЕДЕЛЬНИК": ["История", "Русский язык", "Математика", "Физика", "Химия", "Английский язык", "ОБЗР", ""],
                "ВТОРНИК":     ["Математика", "История", "Русский язык", "Литература", "Биология", "Физика", "", ""],
                "СРЕДА":       ["Физика", "Математика", "Русский язык", "Обществознание", "Химия", "Английский язык", "", ""],
                "ЧЕТВЕРГ":     ["Индивидуальный проект", "Математика", "Физика", "Литература", "История", "Обществознание", "", ""],
                "ПЯТНИЦА":     ["Химия", "Математика", "Русский язык", "Биология", "Физика", "Английский язык", "", ""]
            }},
            { name: "10б", room: "205", lessons: {
                "ПОНЕДЕЛЬНИК": ["Математика", "Русский язык", "История", "Физика", "Химия", "Английский язык", "ОБЗР", ""],
                "ВТОРНИК":     ["Русский язык", "Математика", "Литература", "История", "Биология", "Физика", "", ""],
                "СРЕДА":       ["Физика", "Химия", "Математика", "Русский язык", "История", "Английский язык", "", ""],
                "ЧЕТВЕРГ":     ["Математика", "Русский язык", "Физика", "История", "Обществознание", "Литература", "", ""],
                "ПЯТНИЦА":     ["Химия", "Биология", "Математика", "Русский язык", "Физика", "Английский язык", "", ""]
            }},
            { name: "10в", room: "218", lessons: {
                "ПОНЕДЕЛЬНИК": ["Математика", "Русский язык", "История", "Физика", "Химия", "Биология", "ОБЗР", ""],
                "ВТОРНИК":     ["Русский язык", "Математика", "Литература", "История", "Физика", "Английский язык", "", ""],
                "СРЕДА":       ["Химия", "Физика", "Математика", "Русский язык", "История", "Биология", "", ""],
                "ЧЕТВЕРГ":     ["Индивидуальный проект", "Математика", "Физика", "Литература", "Обществознание", "История", "", ""],
                "ПЯТНИЦА":     ["Физика", "Химия", "Математика", "Русский язык", "Биология", "Английский язык", "", ""]
            }},
            { name: "11а", room: "113", lessons: {
                "ПОНЕДЕЛЬНИК": ["Математика", "Физика", "Русский язык", "Литература", "История", "Обществознание", "Английский язык", ""],
                "ВТОРНИК":     ["Физика", "Математика", "Русский язык", "История", "Химия", "Биология", "", ""],
                "СРЕДА":       ["Физика", "Математика", "Русский язык", "Литература", "Обществознание", "История", "", ""],
                "ЧЕТВЕРГ":     ["Математика", "Физика", "Русский язык", "Литература", "История", "Английский язык", "", ""],
                "ПЯТНИЦА":     ["Физика", "Математика", "Русский язык", "История", "ОБЗР", "Литература", "", ""]
            }}
        ];

        const rawIUP = [
            { name: "Назаренко", className: "5б", lessons: {
                "ПОНЕДЕЛЬНИК": ["Английский язык", "Математика", "Занятия с логопедом", "История", "Русский язык", "Музыка", "Занятие с дефектологом", ""],
                "ВТОРНИК":     ["Физическая культура (вакансия)", "Математика", "Занятие с логопедом", "Русский язык", "История", "Литература", "Занятие с психологом", ""],
                "СРЕДА":       ["История", "Занятие с дефектологом", "Математика", "Английский язык", "Русский язык", "ИЗО", "", ""],
                "ЧЕТВЕРГ":     ["География", "Русский язык", "Технология", "Технология", "Математика", "Литература", "", ""],
                "ПЯТНИЦА":     ["Математика", "Английский язык", "Биология", "Русский язык", "Литература", "Физическая культура (вакансия)", "", ""]
            }},
            { name: "Валиулин", className: "5в", lessons: {
                "ПОНЕДЕЛЬНИК": ["Домоводство", "Окружающий и природный мир", "Музыка и движение", "Физическая культура", "Занятие с психологом", "", "", ""],
                "ВТОРНИК":     ["Занятие с психологом", "Изобразительная деятельность", "Математическое представление", "Окружающий и природный мир", "Занятие с дефектологом", "Человек", "", ""],
                "СРЕДА":       ["Занятие с логопедом", "Речь и альтернативная коммуникация", "Музыка и движение", "Занятия с дефектологом", "Физическая культура", "Технология", "", ""],
                "ЧЕТВЕРГ":     ["", "Изобразительная деятельность", "Человек", "Окружающий социальный мир", "Речь и альтернативная коммуникация", "", "", ""],
                "ПЯТНИЦА":     ["Занятия с логопедом", "Математическое представление", "Речь и альтернативная коммуникация", "Окружающий социальный мир", "Физическая культура", "", "", ""]
            }},
            { name: "Ленин", className: "5в", lessons: {
                "ПОНЕДЕЛЬНИК": ["Математика", "Чтение", "Русский язык", "Технология", "Технология", "Физическая культура", "", ""],
                "ВТОРНИК":     ["", "Мир истории", "География", "Природоведение", "Физическая культура", "Основы социальной жизни", "Занятие с логопедом", ""],
                "СРЕДА":       ["Занятие с дефектологом", "Чтение", "Математика", "Русский язык", "Занятие с психологом", "Физическая культура", "", ""],
                "ЧЕТВЕРГ":     ["Русский язык", "Чтение", "География", "Природоведение", "Мир истории", "Математика", "Занятие с логопедом", ""],
                "ПЯТНИЦА":     ["Русский язык", "Чтение", "Математика", "Технология", "Технология", "Технология", "Технология", ""]
            }},
            { name: "Пашкова", className: "6а", lessons: {
                "ПОНЕДЕЛЬНИК": ["Математика", "География", "Русский язык", "Чтение", "Технология", "Технология", "", ""],
                "ВТОРНИК":     ["Математика", "Природоведение", "Физическая культура", "Мир истории", "Русский язык", "Чтение", "Занятие с логопедом", ""],
                "СРЕДА":       ["Технология", "Технология", "География", "Математика", "Русский язык", "Чтение", "", ""],
                "ЧЕТВЕРГ":     ["Занятие с психологом", "Русский язык", "Чтение", "Основы социальной жизни", "Мир истории", "Физическая культура", "", ""],
                "ПЯТНИЦА":     ["Технология", "Технология", "Занятие с логопедом", "Занятие с дефектологом", "Математика", "Основы социальной жизни", "", ""]
            }},
            { name: "Скандаков", className: "9б", lessons: {
                "ПОНЕДЕЛЬНИК": ["Русский язык", "Биология", "Физическая культура", "География", "Математика", "Технология", "Технология", ""],
                "ВТОРНИК":     ["Чтение", "Информатика", "Занятие с дефектологом", "Математика", "История Отечества", "Физическая культура", "Технология", ""],
                "СРЕДА":       ["Вн. Английский", "Русский язык", "Чтение", "География", "Основы социальной жизни", "Занятие с дефектологом", "Занятие с психологом", ""],
                "ЧЕТВЕРГ":     ["Занятие с логопедом", "История Отечества", "Физическая культура", "Технология", "Технология", "Чтение", "", ""],
                "ПЯТНИЦА":     ["Основы социальной жизни", "Технология", "Русский язык", "Математика", "Чтение", "Биология", "Занятие с логопедом", ""]
            }}
        ];

        const DAYS = ["ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА"];
        const DAY_KEYS = { "ПОНЕДЕЛЬНИК": "pn", "ВТОРНИК": "vt", "СРЕДА": "sr", "ЧЕТВЕРГ": "cht", "ПЯТНИЦА": "pt" };
        const DAY_SHORT = { "ПОНЕДЕЛЬНИК": "Пн", "ВТОРНИК": "Вт", "СРЕДА": "Ср", "ЧЕТВЕРГ": "Чт", "ПЯТНИЦА": "Пт" };
        const MAX_LESSONS = 8;

        const DEFAULT_PASSWORD = "sever2";
        const PASSWORD_KEY = "school_password_v1";

        function getCurrentPassword() {
            try {
                const saved = localStorage.getItem(PASSWORD_KEY);
                if (saved && saved.length > 0) return saved;
            } catch (e) {}
            return DEFAULT_PASSWORD;
        }

        function savePassword(newPassword) {
            try {
                localStorage.setItem(PASSWORD_KEY, newPassword);
                return true;
            } catch (e) {
                console.error(e);
                return false;
            }
        }

        function resetPasswordToDefault() {
            try {
                localStorage.removeItem(PASSWORD_KEY);
                return true;
            } catch (e) {
                return false;
            }
        }

        const STORAGE_KEYS = {
            schedule: "school_schedule_v2",
            history: "school_history_v2",
            theme: "school_theme",
            bells: "school_bells_v1",
            myTeacher: "school_my_teacher",
            scheduleType: "school_schedule_type",
            classes: "school_classes_v1",
            iup: "school_iup_v1",
            myClass: "school_my_class",
            myStudent: "school_my_student",
            substitutions: "school_substitutions_v1",
            substitutionsHistory: "school_substitutions_history_v1",
            dictTeachers: "school_dict_teachers_v1",
            dictSubjects: "school_dict_subjects_v1",
            constructor: "school_constructor_v1",
            mobileView: "school_mobile_view_v1",
            updateBackup: "school_update_backup_v1",
            updateDataBackup: "school_update_databackup_v1",
            lastUpdate: "school_last_update_v1"
        };

        let scheduleData = [];
        let historyData = [];
        let currentDayFilter = "all";
        let editMode = false;
        let consoleUnlocked = false;
        let currentSearchQuery = "";
        let currentTeacherFilter = "";
        let showOnlySubstitutes = false;
        let editingTeacherIndex = -1;
        let historyPanelExpanded = false;
        let pendingExcelImport = null;
        let excelWorkbook = null;

        let securityTabUnlockedUntil = 0;
        let dangerTabUnlockedUntil = 0;
        const TAB_CONFIRM_TIMEOUT = 5 * 60 * 1000;

        let currentScheduleType = localStorage.getItem(STORAGE_KEYS.scheduleType) || "teachers";
        let classesData = [];
        let iupData = [];

        let substitutionsJournal = [];
        let substitutionsHistory = [];
        let currentSubstMonth = "";
        let currentSubstSearch = "";
        let substHistoryExpanded = false;
        let editingSubstId = null;

        let dictTeachersManual = [];
        let dictSubjectsManual = [];

        let substViewMode = "month";
        let currentSubstDay = "";
        let substCalendarVisible = false;
        let substCalMonth = "";

        let constructorConfig = null;
        const FACTORY_CONFIG = {
            header: {
                title: "МОУ «Северная СОШ №2»",
                subtitle: "Белгородского муниципального округа Белгородской области",
                note: "Расписание · 2026–2027 учебный год",
                showTitle: true,
                showSubtitle: true,
                showNote: true,
                titleColor: "#0b2a4a"
            },
            tabs: [
                { key: "teachers", label: "Учительское", icon: "👨‍🏫", visible: true },
                { key: "classes",  label: "Детское",     icon: "🎓",    visible: true },
                { key: "iup",      label: "ИУП 5-8",     icon: "📚",    visible: true },
                { key: "substitutions", label: "Журнал замен", icon: "📌", visible: true }
            ],
            theme: {
                accent: "#1a3a6b",
                bgPage: "#eef3f7",
                bgContainer: "#ffffff",
                textPrimary: "#0b2a4a",
                textSecondary: "#3d5a7a",
                border: "#d0ddee",
                substBg: "#fff3cd",
                substText: "#7a5500",
                success: "#2a6b3a",
                danger: "#b00020"
            },
            sizes: {
                fontSizeBase: 0.78,
                fontSizeTitle: 1.35,
                radius: 12,
                containerWidth: 1800,
                containerPadding: 20,
                rowPadding: 6
            },
            effects: {
                animatedBg: true,
                shadow: true,
                blur: true,
                transitions: true,
                bgIntensity: 1.0
            },
            buttons: {
                edit: true,
                mySchedule: true,
                console: true,
                theme: true,
                mobile: true,
                indicators: true,
                adminIndicator: true
            },
            footer: {
                hint: "Просмотр расписания. Для редактирования нажмите «Редактировать».",
                showStatus: true,
                showHistory: true
            },
            customTabs: []
        };

        let updateTabUnlockedUntil = 0;
        let pendingUpdateHtml = null;
        let pendingUpdateMeta = null;

        function persistClasses() {
            try { localStorage.setItem(STORAGE_KEYS.classes, JSON.stringify(classesData)); } catch (e) { console.error("Ошибка сохранения классов:", e); }
        }

        function persistIUP() {
            try { localStorage.setItem(STORAGE_KEYS.iup, JSON.stringify(iupData)); } catch (e) { console.error("Ошибка сохранения ИУП:", e); }
        }

        function persistSubstitutions() {
            try { localStorage.setItem(STORAGE_KEYS.substitutions, JSON.stringify(substitutionsJournal)); } catch (e) {}
        }

        function persistSubstitutionsHistory() {
            try { localStorage.setItem(STORAGE_KEYS.substitutionsHistory, JSON.stringify(substitutionsHistory)); } catch (e) {}
        }

        function persistDictTeachers() {
            try { localStorage.setItem(STORAGE_KEYS.dictTeachers, JSON.stringify(dictTeachersManual)); } catch (e) {}
        }

        function persistDictSubjects() {
            try { localStorage.setItem(STORAGE_KEYS.dictSubjects, JSON.stringify(dictSubjectsManual)); } catch (e) {}
        }

        function loadDictTeachers() {
            try {
                const saved = localStorage.getItem(STORAGE_KEYS.dictTeachers);
                if (saved) {
                    const arr = JSON.parse(saved);
                    if (Array.isArray(arr)) return arr;
                }
            } catch (e) {}
            return [];
        }

        function loadDictSubjects() {
            try {
                const saved = localStorage.getItem(STORAGE_KEYS.dictSubjects);
                if (saved) {
                    const arr = JSON.parse(saved);
                    if (Array.isArray(arr)) return arr;
                }
            } catch (e) {}
            return [];
        }

        function loadSubstitutions() {
            try {
                const saved = localStorage.getItem(STORAGE_KEYS.substitutions);
                if (saved) {
                    const parsed = JSON.parse(saved);
                    if (Array.isArray(parsed)) {
                        return parsed.map(r => ({
                            id: r.id || ("subst_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8)),
                            date: r.date || new Date().toISOString().slice(0, 10),
                            day: r.day || "",
                            lessonType: r.lessonType || "Стандартный урок",
                            absentTeacher: r.replacedTeacher || r.absentTeacher || "",
                            absentSubject: r.subject || r.absentSubject || "",
                            substituteTeacher: r.substituteTeacher || "",
                            substituteSubject: r.substituteSubject || "",
                            note: r.note || "",
                            status: r.status || "confirmed",
                            createdAt: r.createdAt || new Date().toISOString()
                        }));
                    }
                }
            } catch (e) {}
            return [];
        }

        function loadSubstitutionsHistory() {
            try {
                const saved = localStorage.getItem(STORAGE_KEYS.substitutionsHistory);
                if (saved) {
                    const parsed = JSON.parse(saved);
                    if (Array.isArray(parsed)) return parsed;
                }
            } catch (e) {}
            return [];
        }

        function loadClasses() {
            try {
                const saved = localStorage.getItem(STORAGE_KEYS.classes);
                if (saved) {
                    const parsed = JSON.parse(saved);
                    if (Array.isArray(parsed) && parsed.length) {
                        parsed.forEach(c => {
                            DAYS.forEach(day => {
                                if (!c.lessons[day]) c.lessons[day] = [];
                                for (let i = 0; i < MAX_LESSONS; i++) {
                                    c.lessons[day][i] = normalizeLesson(c.lessons[day][i]);
                                }
                                c.lessons[day] = c.lessons[day].slice(0, MAX_LESSONS);
                                while (c.lessons[day].length < MAX_LESSONS) {
                                    c.lessons[day].push({ text: "", substitute: "" });
                                }
                            });
                        });
                        return parsed;
                    }
                }
            } catch (e) { console.error(e); }
            return null;
        }

        function loadIUP() {
            try {
                const saved = localStorage.getItem(STORAGE_KEYS.iup);
                if (saved) {
                    const parsed = JSON.parse(saved);
                    if (Array.isArray(parsed) && parsed.length) {
                        parsed.forEach(s => {
                            DAYS.forEach(day => {
                                if (!s.lessons[day]) s.lessons[day] = [];
                                for (let i = 0; i < MAX_LESSONS; i++) {
                                    s.lessons[day][i] = normalizeLesson(s.lessons[day][i]);
                                }
                                s.lessons[day] = s.lessons[day].slice(0, MAX_LESSONS);
                                while (s.lessons[day].length < MAX_LESSONS) {
                                    s.lessons[day].push({ text: "", substitute: "" });
                                }
                            });
                        });
                        return parsed;
                    }
                }
            } catch (e) { console.error(e); }
            return null;
        }

        function initAlternativeSchedules() {
            const savedClasses = loadClasses();
            if (savedClasses) {
                classesData = savedClasses;
            } else {
                classesData = rawClasses.map(c => ({
                    name: c.name,
                    room: c.room || "",
                    lessons: (() => {
                        const lessons = {};
                        DAYS.forEach(day => {
                            const arr = c.lessons[day] || [];
                            const padded = [];
                            for (let i = 0; i < MAX_LESSONS; i++) {
                                padded.push({ text: arr[i] || "", substitute: "" });
                            }
                            lessons[day] = padded;
                        });
                        return lessons;
                    })()
                }));
            }

            const savedIUP = loadIUP();
            if (savedIUP) {
                iupData = savedIUP;
            } else {
                iupData = rawIUP.map(s => ({
                    name: s.name,
                    className: s.className || "",
                    lessons: (() => {
                        const lessons = {};
                        DAYS.forEach(day => {
                            const arr = s.lessons[day] || [];
                            const padded = [];
                            for (let i = 0; i < MAX_LESSONS; i++) {
                                padded.push({ text: arr[i] || "", substitute: "" });
                            }
                            lessons[day] = padded;
                        });
                        return lessons;
                    })()
                }));
            }
        }

        function switchSchedule(type) {
            if (type === currentScheduleType) return;

            if (typeof type === "string" && type.startsWith("custom:")) {
                const tabId = type.split(":")[1];
                switchToCustomTab(tabId);
                return;
            }

            currentScheduleType = type;
            localStorage.setItem(STORAGE_KEYS.scheduleType, type);

            document.querySelectorAll(".schedule-tab").forEach(t => {
                t.classList.toggle("active", t.dataset.schedule === type);
            });

            document.body.classList.remove("mode-teachers", "mode-classes", "mode-iup", "mode-substitutions", "mode-custom");
            document.body.classList.add("mode-" + type);

            const hintText = document.getElementById("scheduleHintText");
            if (hintText) {
                if (type === "teachers") {
                    hintText.innerHTML = "<strong>Раздел администратора</strong>: редактирование, замены, консоль, история изменений.";
                } else if (type === "classes") {
                    hintText.innerHTML = "<strong>Расписание по классам</strong>: редактирование предметов и замен.";
                } else if (type === "iup") {
                    hintText.innerHTML = "<strong>Индивидуальные учебные планы</strong>: редактирование занятий и замен.";
                } else if (type === "substitutions") {
                    hintText.innerHTML = "<strong>Журнал замен учителей</strong>: дашборд с формой добавления и статистикой.";
                }
            }

            const searchInput = document.getElementById("teacherSearch");
            if (searchInput) {
                if (type === "teachers") searchInput.placeholder = "Имя учителя...";
                else if (type === "classes") searchInput.placeholder = "Класс (например, 5а)...";
                else searchInput.placeholder = "Имя ученика...";
            }

            currentSearchQuery = "";
            currentTeacherFilter = "";
            showOnlySubstitutes = false;
            const searchInput2 = document.getElementById("teacherSearch");
            const clearBtn = document.getElementById("clearSearch");
            if (searchInput2) searchInput2.value = "";
            if (clearBtn) clearBtn.classList.remove("visible");
            const substFilter = document.getElementById("substituteFilter");
            if (substFilter) substFilter.checked = false;

            populateTeacherSelectForType();

            if (type === "substitutions") {
                renderSubstitutionsPanel();
            } else {
                renderTable();
            }
            updateSubstituteIndicator();
        }

        function getCurrentScheduleData() {
            if (currentScheduleType === "classes") return classesData;
            if (currentScheduleType === "iup") return iupData;
            if (currentScheduleType === "substitutions") return substitutionsJournal;
            return scheduleData;
        }

        function getCurrentFilteredItems() {
            const data = getCurrentScheduleData();
            const query = currentSearchQuery.trim().toLowerCase();
            const selected = currentTeacherFilter;

            return data
                .map((t, idx) => ({ ...t, _idx: idx }))
                .filter(t => {
                    if (selected && t.name !== selected) return false;
                    if (query && !t.name.toLowerCase().includes(query)) return false;
                    if (showOnlySubstitutes && !teacherHasSubstitutes(t)) return false;
                    return true;
                });
        }

        function populateTeacherSelectForType() {
            const select = document.getElementById("teacherSelect");
            const data = getCurrentScheduleData();

            let placeholder = "— Все учителя —";
            if (currentScheduleType === "classes") placeholder = "— Все классы —";
            if (currentScheduleType === "iup") placeholder = "— Все ученики —";

            select.innerHTML = `<option value="">${placeholder}</option>`;
            const names = data.map(t => t.name).sort((a, b) => a.localeCompare(b, "ru"));
            names.forEach(name => {
                const opt = document.createElement("option");
                opt.value = name;
                opt.textContent = name;
                select.appendChild(opt);
            });

            if (currentScheduleType === "teachers") {
                const bulkSel = document.getElementById("bulkTeacherSelect");
                if (bulkSel) {
                    bulkSel.innerHTML = "";
                    names.forEach(name => {
                        const opt = document.createElement("option");
                        opt.value = name;
                        opt.textContent = name;
                        bulkSel.appendChild(opt);
                    });
                }
            }
        }

        function initTheme() {
            const savedTheme = localStorage.getItem(STORAGE_KEYS.theme);
            if (savedTheme === "dark") document.body.classList.add("dark-theme");
            document.getElementById("themeToggle").addEventListener("click", () => {
                document.body.classList.toggle("dark-theme");
                const isDark = document.body.classList.contains("dark-theme");
                localStorage.setItem(STORAGE_KEYS.theme, isDark ? "dark" : "light");
            });
        }

        function isMobileViewEnabled() {
            return localStorage.getItem(STORAGE_KEYS.mobileView) === "1";
        }

        function applyMobileView(enabled) {
            const body = document.body;
            const html = document.documentElement;
            const btn = document.getElementById("mobileToggle");
            const icon = document.getElementById("mobileToggleIcon");
            const label = document.getElementById("mobileToggleLabel");

            if (enabled) {
                body.classList.add("mobile-view");
                html.classList.add("mobile-view");
                if (icon) icon.textContent = "🖥";
                if (label) label.textContent = "Обычная";
                if (btn) {
                    btn.classList.add("active");
                    btn.title = "Вернуться к обычному виду";
                }
            } else {
                body.classList.remove("mobile-view");
                html.classList.remove("mobile-view");
                if (icon) icon.textContent = "📱";
                if (label) label.textContent = "Мобильная";
                if (btn) {
                    btn.classList.remove("active");
                    btn.title = "Переключить мобильную версию";
                }
            }
        }

        function initMobileView() {
            const saved = isMobileViewEnabled();
            applyMobileView(saved);

            const btn = document.getElementById("mobileToggle");
            if (!btn) return;

            btn.addEventListener("click", () => {
                const enabled = !document.body.classList.contains("mobile-view");
                applyMobileView(enabled);
                try {
                    localStorage.setItem(STORAGE_KEYS.mobileView, enabled ? "1" : "0");
                } catch (e) {}

                if (enabled) {
                    showToast("📱 Мобильная версия включена", "success");
                } else {
                    showToast("🖥 Обычная версия", "success");
                }

                if (typeof renderTable === "function") renderTable();
            });
        }

        function loadHistory() {
            try {
                const saved = localStorage.getItem(STORAGE_KEYS.history);
                if (saved) historyData = JSON.parse(saved);
                if (!Array.isArray(historyData)) historyData = [];
            } catch (e) { historyData = []; }
        }

        function saveHistory() {
            try { localStorage.setItem(STORAGE_KEYS.history, JSON.stringify(historyData)); } catch (e) {}
        }

        function addHistoryEntry(entry) {
            const record = {
                id: Date.now() + "_" + Math.random().toString(36).slice(2, 8),
                timestamp: new Date().toISOString(),
                undone: false,
                ...entry
            };
            historyData.unshift(record);
            if (historyData.length > 500) historyData = historyData.slice(0, 500);
            saveHistory();
            updateHistoryUI();
        }

        function undoHistoryEntry(id) {
            const idx = historyData.findIndex(h => h.id === id);
            if (idx === -1) return;
            const entry = historyData[idx];
            if (entry.undone) { showToast("Это изменение уже отменено", "warning"); return; }

            if (entry.type === "edit") {
                let targetArray = scheduleData;
                if (entry.scheduleType === "classes") targetArray = classesData;
                else if (entry.scheduleType === "iup") targetArray = iupData;

                const t = targetArray[entry.teacherIndex];
                if (t && t.lessons[entry.day]) {
                    t.lessons[entry.day][entry.lessonIndex] = {
                        text: entry.oldValue || "",
                        substitute: entry.oldSubstitute || ""
                    };

                    if (entry.scheduleType === "classes") persistClasses();
                    else if (entry.scheduleType === "iup") persistIUP();
                    else persistSchedule();
                }
            } else if (entry.type === "addTeacher") {
                scheduleData = scheduleData.filter((_, i) => i !== entry.teacherIndex);
            } else if (entry.type === "removeTeacher") {
                scheduleData.splice(entry.teacherIndex, 0, entry.teacherSnapshot);
            } else if (entry.type === "bulk" || entry.type === "substituteBulk" || entry.type === "excelImport") {
                if (entry.snapshot) scheduleData = JSON.parse(entry.snapshot);
            } else if (entry.type === "teacherMeta") {
                const t = scheduleData[entry.teacherIndex];
                if (t) {
                    if (entry.field === "name") t.name = entry.oldValue;
                    if (entry.field === "room") t.room = entry.oldValue;
                    if (entry.field === "cls") t.cls = entry.oldValue;
                }
            }

            entry.undone = true;
            saveHistory();
            renderTable();
            populateTeacherSelectForType();
            renderAdminTable();
            updateHistoryUI();
            updateSubstituteIndicator();
            showToast("✓ Изменение отменено", "success");
        }

        function clearHistory() {
            if (!confirm("Очистить всю историю изменений?")) return;
            historyData = [];
            saveHistory();
            updateHistoryUI();
            showToast("История очищена", "success");
        }

        function undoLastChange() {
            const last = historyData.find(h => !h.undone);
            if (!last) { showToast("Нет изменений для отмены", "warning"); return; }
            undoHistoryEntry(last.id);
        }

        function formatTime(iso) {
            try {
                const d = new Date(iso);
                const now = new Date();
                const diffMs = now - d;
                const diffMin = Math.floor(diffMs / 60000);
                if (diffMin < 1) return "только что";
                if (diffMin < 60) return `${diffMin} мин назад`;
                if (diffMin < 60 * 24) return `${Math.floor(diffMin / 60)} ч назад`;
                return d.toLocaleString("ru-RU", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" });
            } catch (e) { return iso; }
        }

        function updateHistoryUI() {
            const count = historyData.filter(h => !h.undone).length;
            const badge = document.getElementById("historyCount");
            if (badge) badge.textContent = count;

            const consoleBadge = document.getElementById("consoleBadge");
            if (consoleBadge) {
                const substCount = countSubstitutes();
                const total = count + substCount;
                if (total > 0) {
                    consoleBadge.textContent = total > 99 ? "99+" : total;
                    consoleBadge.classList.remove("hidden");
                    if (substCount > 0 && count === 0) consoleBadge.classList.add("subst");
                    else consoleBadge.classList.remove("subst");
                } else {
                    consoleBadge.classList.add("hidden");
                }
            }

            renderHistoryList("historyPanelList");
            renderHistoryList("consoleHistoryList");
        }

        function renderHistoryList(containerId) {
            const container = document.getElementById(containerId);
            if (!container) return;

            if (historyData.length === 0) {
                container.innerHTML = '<div class="history-empty">📭 История изменений пуста</div>';
                return;
            }

            container.innerHTML = "";
            historyData.slice(0, 100).forEach(entry => {
                const item = document.createElement("div");
                item.className = "history-item" + (entry.undone ? " undone" : "");
                if (entry.type === "edit" && entry.newSubstitute) item.classList.add("substitute-entry");

                const time = document.createElement("div");
                time.className = "history-time";
                time.textContent = formatTime(entry.timestamp);

                const content = document.createElement("div");
                content.className = "history-content";

                let title = "";
                let detail = "";

                if (entry.type === "edit") {
                    let sectionLabel = "";
                    if (entry.scheduleLabel === "Детское") sectionLabel = ` <span style="font-size:0.65rem; color:var(--text-muted);">[🎓 Детское]</span>`;
                    else if (entry.scheduleLabel === "ИУП 5-8") sectionLabel = ` <span style="font-size:0.65rem; color:var(--text-muted);">[📚 ИУП]</span>`;
                    title = `<span class="who">${escapeHtml(entry.teacherName)}</span> · ${entry.day} · урок ${entry.lessonIndex + 1}${sectionLabel}`;
                    let oldStr = escapeHtml(entry.oldValue || "(пусто)");
                    let newStr = escapeHtml(entry.newValue || "(пусто)");
                    if (entry.oldSubstitute) oldStr += ` <span class="subst-val">(замена: ${escapeHtml(entry.oldSubstitute)})</span>`;
                    if (entry.newSubstitute) newStr += ` <span class="subst-val">(замена: ${escapeHtml(entry.newSubstitute)})</span>`;
                    detail = `<span class="old-val">${oldStr}</span> → <span class="new-val">${newStr}</span>`;
                } else if (entry.type === "addTeacher") {
                    title = `➕ Добавлен учитель <span class="who">${escapeHtml(entry.teacherName)}</span>`;
                    detail = `Каб.: ${escapeHtml(entry.room || "—")}, Кл.: ${escapeHtml(entry.cls || "—")}`;
                } else if (entry.type === "removeTeacher") {
                    title = `🗑 Удалён учитель <span class="who">${escapeHtml(entry.teacherName)}</span>`;
                    detail = `Все данные сохранены в истории`;
                } else if (entry.type === "teacherMeta") {
                    title = `✎ Изменено поле «${entry.field}» у <span class="who">${escapeHtml(entry.teacherName)}</span>`;
                    detail = `<span class="old-val">${escapeHtml(entry.oldValue || "—")}</span> → <span class="new-val">${escapeHtml(entry.newValue || "—")}</span>`;
                } else if (entry.type === "bulk" || entry.type === "substituteBulk") {
                    title = `⚡ ${escapeHtml(entry.description || "Массовая операция")}`;
                    detail = `Затронуто записей: ${entry.affected || "—"}`;
                } else if (entry.type === "excelImport") {
                    title = `📊 Импорт из Excel: <span class="who">${escapeHtml(entry.fileName || "")}</span>`;
                    detail = `Импортировано учителей: ${entry.affected || 0}, уроков: ${entry.lessons || 0}`;
                } else if (entry.type === "passwordChange") {
                    title = `🔐 ${escapeHtml(entry.description || "Изменение пароля")}`;
                    detail = `Пароль администратора обновлён`;
                }

                content.innerHTML = `<div>${title}</div><div class="detail">${detail}</div>`;

                const actions = document.createElement("div");
                if (!entry.undone) {
                    const undoBtn = document.createElement("button");
                    undoBtn.className = "icon-btn";
                    undoBtn.title = "Откатить";
                    undoBtn.textContent = "↶";
                    undoBtn.addEventListener("click", () => undoHistoryEntry(entry.id));
                    actions.appendChild(undoBtn);
                } else {
                    actions.innerHTML = '<span style="color:var(--text-muted); font-size:0.7rem;">отменено</span>';
                }

                item.appendChild(time);
                item.appendChild(content);
                item.appendChild(actions);
                container.appendChild(item);
            });
        }

        function escapeHtml(str) {
            if (str == null) return "";
            return String(str)
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#39;");
        }

        function normalizeLesson(val) {
            if (val === null || val === undefined) return { text: "", substitute: "" };
            if (typeof val === "string") return { text: val, substitute: "" };
            if (typeof val === "object") {
                return { text: val.text || "", substitute: val.substitute || "" };
            }
            return { text: "", substitute: "" };
        }

        function buildScheduleFromRaw() {
            return rawTeachers.map(t => {
                const lessons = {};
                DAYS.forEach(day => {
                    const key = DAY_KEYS[day];
                    const arr = t[key] || [];
                    const padded = [];
                    for (let i = 0; i < MAX_LESSONS; i++) {
                        padded.push({ text: arr[i] || "", substitute: "" });
                    }
                    lessons[day] = padded;
                });
                return { name: t.name, room: t.room || "", cls: t.cls || "", lessons };
            });
        }

        function loadData() {
            const saved = localStorage.getItem(STORAGE_KEYS.schedule);
            if (saved) {
                try {
                    const parsed = JSON.parse(saved);
                    if (Array.isArray(parsed) && parsed.length) {
                        parsed.forEach(t => {
                            DAYS.forEach(day => {
                                if (!t.lessons[day]) t.lessons[day] = [];
                                for (let i = 0; i < MAX_LESSONS; i++) {
                                    t.lessons[day][i] = normalizeLesson(t.lessons[day][i]);
                                }
                                t.lessons[day] = t.lessons[day].slice(0, MAX_LESSONS);
                                while (t.lessons[day].length < MAX_LESSONS) {
                                    t.lessons[day].push({ text: "", substitute: "" });
                                }
                            });
                        });
                        scheduleData = parsed;
                        return;
                    }
                } catch (e) { /* ignore */ }
            }
            scheduleData = buildScheduleFromRaw();
        }

        function persistSchedule() {
            localStorage.setItem(STORAGE_KEYS.schedule, JSON.stringify(scheduleData));
        }

        function saveData() {
            if (currentScheduleType === "teachers") persistSchedule();
            else if (currentScheduleType === "classes") persistClasses();
            else if (currentScheduleType === "iup") persistIUP();
            showStatus("Сохранено", "var(--success)");
            showToast("💾 Расписание сохранено", "success");
        }

        function resetData() {
            if (!confirm("Сбросить все изменения и вернуть исходное расписание?")) return;
            if (currentScheduleType === "teachers") {
                const snapshot = JSON.stringify(scheduleData);
                scheduleData = buildScheduleFromRaw();
                addHistoryEntry({ type: "bulk", description: "Полный сброс к исходным данным", snapshot, affected: scheduleData.length, scheduleType: "teachers" });
                persistSchedule();
            } else if (currentScheduleType === "classes") {
                classesData = rawClasses.map(c => ({
                    name: c.name, room: c.room || "",
                    lessons: (() => {
                        const lessons = {};
                        DAYS.forEach(day => {
                            const arr = c.lessons[day] || [];
                            const padded = [];
                            for (let i = 0; i < MAX_LESSONS; i++) padded.push({ text: arr[i] || "", substitute: "" });
                            lessons[day] = padded;
                        });
                        return lessons;
                    })()
                }));
                persistClasses();
            } else if (currentScheduleType === "iup") {
                iupData = rawIUP.map(s => ({
                    name: s.name, className: s.className || "",
                    lessons: (() => {
                        const lessons = {};
                        DAYS.forEach(day => {
                            const arr = s.lessons[day] || [];
                            const padded = [];
                            for (let i = 0; i < MAX_LESSONS; i++) padded.push({ text: arr[i] || "", substitute: "" });
                            lessons[day] = padded;
                        });
                        return lessons;
                    })()
                }));
                persistIUP();
            }
            renderTable();
            populateTeacherSelectForType();
            updateSubstituteIndicator();
            showStatus("Сброшено", "var(--warning)");
            showToast("↺ Расписание сброшено к исходному", "warning");
        }

        function showStatus(msg, color) {
            const badge = document.getElementById("statusBadge");
            if (badge) {
                badge.textContent = msg;
                badge.style.background = color || "var(--success)";
                setTimeout(() => {
                    badge.textContent = "Сохранено локально";
                    badge.style.background = "var(--success)";
                }, 1500);
            }
        }

        function showToast(msg, type = "success") {
            const existing = document.querySelector(".toast");
            if (existing) existing.remove();
            const toast = document.createElement("div");
            toast.className = "toast" + (type === "error" ? " error" : type === "warning" ? " warning" : type === "substitute" ? " substitute" : "");
            toast.textContent = msg;
            document.body.appendChild(toast);
            setTimeout(() => toast.remove(), 2800);
        }

        function countSubstitutes() {
            let count = 0;
            const data = [scheduleData, classesData, iupData];
            data.forEach(arr => {
                arr.forEach(t => {
                    DAYS.forEach(day => {
                        (t.lessons[day] || []).forEach(l => {
                            if (l && l.substitute && l.substitute.trim()) count++;
                        });
                    });
                });
            });
            return count;
        }

        function countTotalLessons() {
            let count = 0;
            scheduleData.forEach(t => {
                DAYS.forEach(day => {
                    (t.lessons[day] || []).forEach(l => {
                        const lesson = normalizeLesson(l);
                        if (lesson.text && lesson.text.trim()) count++;
                    });
                });
            });
            return count;
        }

        function updateSubstituteIndicator() {
            const count = countSubstitutes();
            const indicator = document.getElementById("substituteIndicator");
            const countEl = document.getElementById("substituteIndicatorCount");
            if (countEl) countEl.textContent = count;
            if (indicator) {
                if (count > 0) indicator.classList.add("visible");
                else indicator.classList.remove("visible");
            }
            updateHistoryUI();
        }

        function getFreeTeachers(day, lessonIndex, excludeTeacherName) {
            const free = [];
            scheduleData.forEach((t, idx) => {
                if (t.name === excludeTeacherName) return;
                const lesson = normalizeLesson(t.lessons[day]?.[lessonIndex]);
                const hasOwnLesson = lesson.text && lesson.text.trim();
                const hasSubstituteWork = lesson.substitute && lesson.substitute.trim();
                if (!hasOwnLesson && !hasSubstituteWork) {
                    free.push({ name: t.name, room: t.room || "", cls: t.cls || "", index: idx });
                }
            });
            return free.sort((a, b) => a.name.localeCompare(b.name, "ru"));
        }

        function teacherHasSubstitutes(t) {
            for (const day of DAYS) {
                for (const l of (t.lessons[day] || [])) {
                    if (l && l.substitute && l.substitute.trim()) return true;
                }
            }
            return false;
        }

        function getFilteredTeachers() {
            return getCurrentFilteredItems();
        }

        function updateFilterCount(count) {
            const el = document.getElementById("filterCount");
            if (!el) return;
            const total = getCurrentScheduleData().length;
            if (count === total && !currentSearchQuery && !currentTeacherFilter && !showOnlySubstitutes) {
                el.textContent = `Всего: ${total}`;
            } else {
                el.textContent = `Найдено: ${count} из ${total}`;
            }
        }

        function populateTeacherSelect() {
            populateTeacherSelectForType();
        }

        function initFilters() {
            const searchInput = document.getElementById("teacherSearch");
            const clearBtn = document.getElementById("clearSearch");
            const select = document.getElementById("teacherSelect");
            const substFilter = document.getElementById("substituteFilter");

            searchInput.addEventListener("input", (e) => {
                currentSearchQuery = e.target.value;
                clearBtn.classList.toggle("visible", currentSearchQuery.length > 0);
                if (currentSearchQuery) { select.value = ""; currentTeacherFilter = ""; }
                renderTable();
            });

            clearBtn.addEventListener("click", () => {
                searchInput.value = "";
                currentSearchQuery = "";
                clearBtn.classList.remove("visible");
                renderTable();
                searchInput.focus();
            });

            select.addEventListener("change", (e) => {
                currentTeacherFilter = e.target.value;
                if (currentTeacherFilter) {
                    searchInput.value = "";
                    currentSearchQuery = "";
                    clearBtn.classList.remove("visible");
                }
                renderTable();
            });

            substFilter.addEventListener("change", (e) => {
                showOnlySubstitutes = e.target.checked;
                renderTable();
            });
        }

        function renderTable() {
            if (currentScheduleType === "substitutions") {
                renderSubstitutionsPanel();
                return;
            }

            const daysToShow = currentDayFilter === "all" ? DAYS : [currentDayFilter];

            const thead = document.getElementById("scheduleThead");
            thead.innerHTML = "";

            const row1 = document.createElement("tr");
            const cornerTh = document.createElement("th");
            let firstColLabel = "Учитель";
            if (currentScheduleType === "classes") firstColLabel = "Класс";
            if (currentScheduleType === "iup") firstColLabel = "Ученик";
            cornerTh.textContent = firstColLabel;
            cornerTh.style.background = "var(--th-bg-darker)";
            cornerTh.rowSpan = 2;
            cornerTh.style.verticalAlign = "middle";
            cornerTh.style.minWidth = "140px";
            cornerTh.style.position = "sticky";
            cornerTh.style.left = "0";
            cornerTh.style.zIndex = "7";
            row1.appendChild(cornerTh);

            const roomTh = document.createElement("th");
            roomTh.textContent = currentScheduleType === "iup" ? "Класс" : "Каб.";
            roomTh.rowSpan = 2;
            roomTh.style.background = "var(--th-bg-darker)";
            roomTh.style.verticalAlign = "middle";
            roomTh.style.minWidth = "48px";
            roomTh.style.position = "sticky";
            roomTh.style.left = "140px";
            roomTh.style.zIndex = "7";
            row1.appendChild(roomTh);

            const clsTh = document.createElement("th");
            if (currentScheduleType === "iup") {
                clsTh.style.display = "none";
            }
            clsTh.textContent = "Кл.";
            clsTh.rowSpan = 2;
            clsTh.style.background = "var(--th-bg-darker)";
            clsTh.style.verticalAlign = "middle";
            clsTh.style.minWidth = "44px";
            clsTh.style.position = "sticky";
            clsTh.style.left = "188px";
            clsTh.style.zIndex = "7";
            row1.appendChild(clsTh);

            daysToShow.forEach(day => {
                const th = document.createElement("th");
                th.colSpan = MAX_LESSONS;
                th.textContent = day;
                th.style.background = "var(--th-bg-dark)";
                th.style.fontSize = "0.78rem";
                row1.appendChild(th);
            });
            thead.appendChild(row1);

            const row2 = document.createElement("tr");
            daysToShow.forEach(() => {
                for (let lesson = 1; lesson <= MAX_LESSONS; lesson++) {
                    const th = document.createElement("th");
                    th.textContent = lesson;
                    th.style.background = "var(--th-bg)";
                    th.style.fontSize = "0.68rem";
                    th.style.padding = "4px 2px";
                    row2.appendChild(th);
                }
            });
            thead.appendChild(row2);

            const tbody = document.getElementById("scheduleBody");
            tbody.innerHTML = "";

            const filtered = getCurrentFilteredItems();
            updateFilterCount(filtered.length);
            updateSubstituteIndicator();

            if (filtered.length === 0) {
                const tr = document.createElement("tr");
                tr.className = "no-data-row";
                const td = document.createElement("td");
                td.colSpan = 3 + daysToShow.length * MAX_LESSONS;
                if (currentScheduleType === "classes") {
                    td.textContent = "🔍 Классы по заданным условиям не найдены";
                } else if (currentScheduleType === "iup") {
                    td.textContent = "🔍 Ученики по заданным условиям не найдены";
                } else {
                    td.textContent = showOnlySubstitutes
                        ? "📌 В текущем виде нет ячеек с заменами"
                        : "🔍 Учителя по заданным условиям не найдены";
                }
                tr.appendChild(td);
                tbody.appendChild(tr);
                return;
            }

            filtered.forEach((item) => {
                const tIndex = item._idx;
                const tr = document.createElement("tr");

                const tdName = document.createElement("td");
                tdName.className = "teacher-name";
                tdName.textContent = item.name;
                tr.appendChild(tdName);

                const tdRoom = document.createElement("td");
                tdRoom.className = "room-cell";
                if (currentScheduleType === "iup") {
                    tdRoom.textContent = item.className || "—";
                } else {
                    tdRoom.textContent = item.room || "—";
                }
                tr.appendChild(tdRoom);

                const tdCls = document.createElement("td");
                tdCls.className = "class-cell";
                if (currentScheduleType === "iup") {
                    tdCls.style.display = "none";
                }
                tdCls.textContent = item.cls || "—";
                tr.appendChild(tdCls);

                daysToShow.forEach(day => {
                    const lessons = item.lessons[day] || [];
                    for (let i = 0; i < MAX_LESSONS; i++) {
                        const lesson = normalizeLesson(lessons[i]);
                        const td = document.createElement("td");
                        td.className = "lesson-cell";
                        td.dataset.teacherIndex = tIndex;
                        td.dataset.day = day;
                        td.dataset.lessonIndex = i;
                        td.dataset.lessonNum = i + 1;

                        renderCellContent(td, lesson);

                        td.addEventListener("click", function() {
                            if (!editMode) return;
                            if (this.classList.contains("editing")) return;
                            startEdit(this);
                        });

                        tr.appendChild(td);
                    }
                });

                tbody.appendChild(tr);
            });
        }

        function renderCellContent(cell, lesson) {
            cell.innerHTML = "";
            cell.classList.remove("has-substitute", "empty-lesson");

            const hasSubst = lesson.substitute && lesson.substitute.trim();
            const hasText = lesson.text && lesson.text.trim();

            if (!hasText && !hasSubst) {
                cell.classList.add("empty-lesson");
                return;
            }

            if (hasSubst) {
                cell.classList.add("has-substitute");
                const icon = document.createElement("span");
                icon.className = "substitute-icon";
                icon.textContent = "📌";
                cell.appendChild(icon);

                if (hasText) {
                    const textSpan = document.createElement("span");
                    textSpan.className = "lesson-text";
                    textSpan.textContent = lesson.text;
                    cell.appendChild(textSpan);
                }

                const note = document.createElement("span");
                note.className = "substitute-note";
                if (currentScheduleType === "teachers") {
                    note.textContent = `(замена у ${lesson.substitute})`;
                } else {
                    note.textContent = `(замена: ${lesson.substitute})`;
                }
                cell.appendChild(note);
            } else if (hasText) {
                cell.textContent = lesson.text;
            }
        }

        function startEdit(td) {
            if (!editMode) return;

            const itemIndex = parseInt(td.dataset.teacherIndex);
            const day = td.dataset.day;
            const lessonIndex = parseInt(td.dataset.lessonIndex);

            const dataArray = getCurrentScheduleData();
            const currentLesson = normalizeLesson(dataArray[itemIndex].lessons[day][lessonIndex]);
            const currentTeacherName = dataArray[itemIndex].name;

            td.classList.add("editing");
            td.innerHTML = "";

            const editor = document.createElement("div");
            editor.className = "cell-editor";

            const textInput = document.createElement("input");
            textInput.type = "text";
            textInput.value = currentLesson.text;
            textInput.placeholder = "Предмет...";
            editor.appendChild(textInput);

            const substLabel = document.createElement("div");
            substLabel.className = "substitute-label";
            substLabel.innerHTML = `<span>📌 Замена</span>`;
            const freeCounter = document.createElement("span");
            freeCounter.className = "free-counter";
            substLabel.appendChild(freeCounter);
            editor.appendChild(substLabel);

            const substRow = document.createElement("div");
            substRow.className = "substitute-row";

            const substInput = document.createElement("input");
            substInput.type = "text";
            substInput.className = "substitute-input";
            substInput.value = currentLesson.substitute;
            substInput.placeholder = "ФИО учителя (ручной ввод)";
            substInput.autocomplete = "off";
            substRow.appendChild(substInput);

            const clearBtn = document.createElement("button");
            clearBtn.className = "clear-substitute";
            clearBtn.type = "button";
            clearBtn.textContent = "✕";
            clearBtn.title = "Очистить замену";
            clearBtn.addEventListener("click", (e) => {
                e.stopPropagation();
                substInput.value = "";
                substInput.focus();
            });
            substRow.appendChild(clearBtn);

            editor.appendChild(substRow);

            const hint = document.createElement("div");
            hint.className = "editor-hint";
            hint.textContent = "Enter — сохранить, Esc — отмена";
            editor.appendChild(hint);

            td.appendChild(editor);

            let freeTeachers = [];
            let dropdown = null;

            if (currentScheduleType === "teachers") {
                dropdown = document.createElement("div");
                dropdown.className = "free-teachers-dropdown hidden";
                substRow.appendChild(dropdown);

                function refreshFreeTeachers() {
                    freeTeachers = getFreeTeachers(day, lessonIndex, currentTeacherName);
                    const total = scheduleData.length - 1;
                    freeCounter.textContent = `Свободно: ${freeTeachers.length} из ${total}`;
                    freeCounter.classList.toggle("all-busy", freeTeachers.length === 0);
                }

                function updateDropdown(filterText) {
                    dropdown.innerHTML = "";
                    const q = (filterText || "").trim().toLowerCase();

                    const header = document.createElement("div");
                    header.className = "dropdown-header";
                    const headerText = document.createElement("span");
                    headerText.textContent = `Свободные (${freeTeachers.length})`;
                    header.appendChild(headerText);

                    const refreshBtn = document.createElement("button");
                    refreshBtn.className = "refresh-btn";
                    refreshBtn.type = "button";
                    refreshBtn.title = "Обновить";
                    refreshBtn.textContent = "🔄";
                    refreshBtn.addEventListener("click", (e) => {
                        e.stopPropagation();
                        refreshFreeTeachers();
                        updateDropdown(substInput.value);
                    });
                    header.appendChild(refreshBtn);
                    dropdown.appendChild(header);

                    const filtered = q
                        ? freeTeachers.filter(t => t.name.toLowerCase().includes(q))
                        : freeTeachers;

                    if (filtered.length === 0) {
                        const empty = document.createElement("div");
                        empty.className = "empty-item";
                        empty.textContent = freeTeachers.length === 0 ? "😕 Все заняты" : "🔍 Никто не найден";
                        dropdown.appendChild(empty);
                        return;
                    }

                    filtered.forEach((t) => {
                        const item = document.createElement("div");
                        item.className = "dropdown-item";

                        const dot = document.createElement("span");
                        dot.className = "status-dot";
                        item.appendChild(dot);

                        const info = document.createElement("div");
                        info.className = "teacher-info";

                        const nameEl = document.createElement("span");
                        nameEl.className = "teacher-name-item";
                        nameEl.textContent = t.name;
                        info.appendChild(nameEl);

                        const metaParts = [];
                        if (t.room) metaParts.push(`каб. ${t.room}`);
                        if (t.cls) metaParts.push(`кл. ${t.cls}`);
                        if (metaParts.length) {
                            const meta = document.createElement("span");
                            meta.className = "teacher-meta";
                            meta.textContent = metaParts.join(" · ");
                            info.appendChild(meta);
                        }

                        item.appendChild(info);

                        item.addEventListener("mousedown", (e) => {
                            e.preventDefault();
                            substInput.value = t.name;
                            dropdown.classList.add("hidden");
                            substInput.focus();
                        });

                        dropdown.appendChild(item);
                    });
                }

                function showDropdown() {
                    refreshFreeTeachers();
                    updateDropdown(substInput.value);
                    dropdown.classList.remove("hidden");
                }

                function hideDropdown() {
                    dropdown.classList.add("hidden");
                }

                substInput.addEventListener("focus", showDropdown);
                substInput.addEventListener("input", () => {
                    updateDropdown(substInput.value);
                    if (dropdown.classList.contains("hidden")) dropdown.classList.remove("hidden");
                });

                const outsideClickHandler = (e) => {
                    if (!substRow.contains(e.target) && e.target !== substInput) hideDropdown();
                };
                document.addEventListener("mousedown", outsideClickHandler);
                td._outsideClickHandler = outsideClickHandler;
            } else {
                freeCounter.textContent = "ручной ввод";
                freeCounter.style.color = "var(--text-muted)";
                substInput.placeholder = "Иванов И.И.";
            }

            let finished = false;
            function finish(save) {
                if (finished) return;
                finished = true;

                if (td._outsideClickHandler) {
                    document.removeEventListener("mousedown", td._outsideClickHandler);
                    delete td._outsideClickHandler;
                }

                const newText = save ? textInput.value.trim() : currentLesson.text;
                const newSubst = save ? substInput.value.trim() : currentLesson.substitute;

                const oldText = currentLesson.text;
                const oldSubst = currentLesson.substitute;
                const changed = (oldText !== newText) || (oldSubst !== newSubst);

                if (changed) {
                    dataArray[itemIndex].lessons[day][lessonIndex] = {
                        text: newText,
                        substitute: newSubst
                    };

                    let scheduleLabel = "Учительское";
                    if (currentScheduleType === "classes") scheduleLabel = "Детское";
                    if (currentScheduleType === "iup") scheduleLabel = "ИУП 5-8";

                    addHistoryEntry({
                        type: "edit",
                        scheduleType: currentScheduleType,
                        scheduleLabel: scheduleLabel,
                        teacherIndex: itemIndex,
                        teacherName: dataArray[itemIndex].name,
                        day,
                        lessonIndex,
                        oldValue: oldText,
                        newValue: newText,
                        oldSubstitute: oldSubst,
                        newSubstitute: newSubst
                    });

                    if (currentScheduleType === "teachers") persistSchedule();
                    else if (currentScheduleType === "classes") persistClasses();
                    else if (currentScheduleType === "iup") persistIUP();
                }

                td.classList.remove("editing");
                renderCellContent(td, { text: newText, substitute: newSubst });

                if (changed) {
                    td.classList.add("recently-changed");
                    setTimeout(() => td.classList.remove("recently-changed"), 1500);
                    if (newSubst && newSubst !== oldSubst) {
                        showToast(`📌 Замена: ${newSubst}`, "substitute");
                    }
                }

                updateSubstituteIndicator();
            }

            textInput.addEventListener("keydown", (e) => {
                if (e.key === "Enter") { e.preventDefault(); finish(true); }
                if (e.key === "Escape") { e.preventDefault(); finish(false); }
                if (e.key === "Tab" && !e.shiftKey) {
                    e.preventDefault();
                    substInput.focus();
                }
            });

            substInput.addEventListener("keydown", (e) => {
                if (e.key === "Enter") {
                    e.preventDefault();
                    if (currentScheduleType === "teachers") {
                        const q = substInput.value.trim().toLowerCase();
                        const matches = q
                            ? freeTeachers.filter(t => t.name.toLowerCase().includes(q))
                            : freeTeachers;
                        if (!substInput.value.trim() && matches.length === 1) {
                            substInput.value = matches[0].name;
                        }
                    }
                    finish(true);
                }
                if (e.key === "Escape") { e.preventDefault(); finish(false); }
                if (e.key === "Tab" && e.shiftKey) {
                    e.preventDefault();
                    textInput.focus();
                }
                if (currentScheduleType === "teachers" && (e.key === "ArrowDown" || e.key === "ArrowUp")) {
                    e.preventDefault();
                    if (!dropdown) return;
                    const items = dropdown.querySelectorAll(".dropdown-item");
                    if (!items.length) return;
                    let currentIdx = -1;
                    items.forEach((it, i) => { if (it.classList.contains("highlighted")) currentIdx = i; });
                    items.forEach(it => it.classList.remove("highlighted"));
                    if (e.key === "ArrowDown") currentIdx = (currentIdx + 1) % items.length;
                    else currentIdx = (currentIdx - 1 + items.length) % items.length;
                    items[currentIdx].classList.add("highlighted");
                    items[currentIdx].scrollIntoView({ block: "nearest" });
                }
            });

            const blurHandler = (e) => {
                if (!td.contains(e.relatedTarget) && e.relatedTarget !== document.body) {
                    setTimeout(() => {
                        if (!finished && !td.contains(document.activeElement)) finish(true);
                    }, 200);
                }
            };
            textInput.addEventListener("blur", blurHandler);
            substInput.addEventListener("blur", blurHandler);

            textInput.focus();
            textInput.select();
        }

        function initDayFilter() {
            const container = document.getElementById("dayFilter");
            container.addEventListener("click", (e) => {
                const btn = e.target.closest(".day-btn");
                if (!btn) return;
                document.querySelectorAll(".day-btn").forEach(b => b.classList.remove("active"));
                btn.classList.add("active");
                currentDayFilter = btn.dataset.day;
                renderTable();
            });
        }

        function enterEditMode() {
            editMode = true;
            document.body.classList.add("edit-mode");

            document.getElementById("editIndicator").classList.add("visible");
            document.getElementById("resetBtn").style.display = "inline-flex";
            document.getElementById("saveBtn").style.display = "inline-flex";

            const btn = document.getElementById("editToggleBtn");
            btn.textContent = "🔒 Заблокировать";
            btn.classList.remove("btn-primary");
            btn.classList.add("btn-warning");

            let hintText = "Режим редактирования активен. Нажмите на любую ячейку, чтобы изменить.";
            if (currentScheduleType === "classes") hintText = "Режим редактирования «Детского» расписания. Замены — с ручным вводом ФИО.";
            if (currentScheduleType === "iup") hintText = "Режим редактирования «ИУП». Замены — с ручным вводом ФИО.";
            document.getElementById("footerHint").textContent = hintText;

            updateAdminModeUI();
        }

        function exitEditMode() {
            editMode = false;
            document.body.classList.remove("edit-mode");

            document.getElementById("editIndicator").classList.remove("visible");
            document.getElementById("resetBtn").style.display = "none";
            document.getElementById("saveBtn").style.display = "none";

            const btn = document.getElementById("editToggleBtn");
            btn.textContent = "✎ Редактировать";
            btn.classList.remove("btn-warning");
            btn.classList.add("btn-primary");

            document.getElementById("footerHint").textContent =
                "Просмотр расписания. Для редактирования нажмите «Редактировать».";

            updateAdminModeUI();
        }

        let passwordCallback = null;

        function openPasswordModal(title, desc, callback) {
            const modal = document.getElementById("passwordModal");
            const input = document.getElementById("passwordInput");
            const error = document.getElementById("passwordError");
            document.getElementById("passwordModalTitle").textContent = title || "🔒 Вход";
            document.getElementById("passwordModalDesc").textContent = desc || "Введите пароль.";
            input.value = ""; error.textContent = "";
            passwordCallback = callback;
            modal.classList.add("visible");
            setTimeout(() => input.focus(), 50);
        }

        function closePasswordModal() {
            document.getElementById("passwordModal").classList.remove("visible");
            passwordCallback = null;
        }

        function tryConfirmPassword() {
            const input = document.getElementById("passwordInput");
            const error = document.getElementById("passwordError");
            if (input.value === getCurrentPassword()) {
                error.textContent = "";
                const cb = passwordCallback;
                closePasswordModal();
                if (cb) cb();
            } else {
                error.textContent = "❌ Неверный пароль.";
                input.value = "";
                input.focus();
            }
        }

        function openConsole() {
            if (!consoleUnlocked) {
                openPasswordModal(
                    "🛠 Доступ к консоли администратора",
                    "Введите пароль для получения доступа.",
                    () => {
                        consoleUnlocked = true;
                        updateAdminModeUI();
                        document.getElementById("consoleModal").classList.add("visible");
                        refreshConsole();
                    }
                );
                return;
            }
            document.getElementById("consoleModal").classList.add("visible");
            refreshConsole();
        }

        function updateAdminModeUI() {
            const indicator = document.getElementById("adminIndicator");
            const logoutBtn = document.getElementById("logoutAdminBtn");
            if (consoleUnlocked || editMode) {
                if (indicator) indicator.classList.add("visible");
                if (logoutBtn) logoutBtn.style.display = "inline-flex";
            } else {
                if (indicator) indicator.classList.remove("visible");
                if (logoutBtn) logoutBtn.style.display = "none";
            }
        }

        function logoutAdmin() {
            if (!confirm("Выйти из режима администратора?\n\nВсе права на редактирование и консоль будут сняты.")) return;

            consoleUnlocked = false;
            editMode = false;
            securityTabUnlockedUntil = 0;
            dangerTabUnlockedUntil = 0;
            constructorTabUnlockedUntil = 0;
            updateTabUnlockedUntil = 0;

            document.body.classList.remove("edit-mode");
            document.getElementById("editIndicator").classList.remove("visible");
            document.getElementById("resetBtn").style.display = "none";
            document.getElementById("saveBtn").style.display = "none";

            const editBtn = document.getElementById("editToggleBtn");
            editBtn.textContent = "✎ Редактировать";
            editBtn.classList.remove("btn-warning");
            editBtn.classList.add("btn-primary");

            document.getElementById("consoleModal").classList.remove("visible");
            document.getElementById("passwordModal").classList.remove("visible");
            document.getElementById("teacherEditModal").classList.remove("visible");
            document.getElementById("myScheduleModal").classList.remove("visible");
            document.getElementById("bellsModal").classList.remove("visible");
            document.getElementById("dictManageModal").classList.remove("visible");

            activateTab("teachers");
            updateAdminModeUI();

            document.getElementById("footerHint").textContent =
                "Просмотр расписания. Для редактирования нажмите «Редактировать».";

            showToast("🚪 Вы вышли из режима администратора", "warning");
        }

        function closeConsole() {
            document.getElementById("consoleModal").classList.remove("visible");
            const opts = document.getElementById("excelImportOptions");
            if (opts) opts.style.display = "none";
            pendingExcelImport = null;
        }

        function refreshConsole() {
            renderAdminTable();
            updateBulkStats();
            updateHistoryUI();
            updateStorageInfo();
            updateCurrentPasswordInfo();
            updateAdminModeUI();
            if (typeof refreshUpdateInfo === "function") refreshUpdateInfo();
        }

        function initConsoleTabs() {
            const tabs = document.querySelectorAll(".console-tab");
            tabs.forEach(tab => {
                if (tab.dataset.dictTab) return;
                tab.addEventListener("click", () => {
                    const tabName = tab.dataset.tab;
                    const requiresPassword = tab.dataset.requiresPassword === "true";

                    if (requiresPassword) {
                        const now = Date.now();
                        let unlockedUntil = 0;
                        if (tabName === "security") unlockedUntil = securityTabUnlockedUntil;
                        else if (tabName === "danger") unlockedUntil = dangerTabUnlockedUntil;
                        else if (tabName === "constructor") unlockedUntil = constructorTabUnlockedUntil;
                        else if (tabName === "update") unlockedUntil = updateTabUnlockedUntil;

                        if (now > unlockedUntil) {
                            openPasswordModal(
                                "🔐 Подтверждение доступа",
                                "Введите пароль администратора. Подтверждение действует 5 минут.",
                                () => {
                                    if (tabName === "security") securityTabUnlockedUntil = Date.now() + TAB_CONFIRM_TIMEOUT;
                                    else if (tabName === "danger") dangerTabUnlockedUntil = Date.now() + TAB_CONFIRM_TIMEOUT;
                                    else if (tabName === "constructor") constructorTabUnlockedUntil = Date.now() + TAB_CONFIRM_TIMEOUT;
                                    else if (tabName === "update") updateTabUnlockedUntil = Date.now() + TAB_CONFIRM_TIMEOUT;
                                    activateTab(tabName);
                                }
                            );
                            return;
                        }
                    }

                    activateTab(tabName);
                });
            });
        }

        function activateTab(tabName) {
            document.querySelectorAll(".console-tab").forEach(t => {
                if (!t.dataset.dictTab) t.classList.remove("active");
            });
            document.querySelectorAll(".console-tab-content").forEach(c => c.classList.remove("active"));

            const tab = document.querySelector(`.console-tab[data-tab="${tabName}"]`);
            const content = document.querySelector(`.console-tab-content[data-tab="${tabName}"]`);

            if (tab) tab.classList.add("active");
            if (content) content.classList.add("active");

            if (tabName === "update" && typeof refreshUpdateInfo === "function") {
                refreshUpdateInfo();
            }
        }

        function renderAdminTable() {
            const tbody = document.getElementById("adminTableBody");
            if (!tbody) return;
            tbody.innerHTML = "";

            const countEl = document.getElementById("adminCount");
            if (countEl) countEl.textContent = `Всего: ${scheduleData.length}`;

            scheduleData.forEach((t, idx) => {
                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td>${idx + 1}</td>
                    <td><input type="text" value="${escapeHtml(t.name)}" data-field="name" data-idx="${idx}"></td>
                    <td><input type="text" value="${escapeHtml(t.room || "")}" data-field="room" data-idx="${idx}" placeholder="—"></td>
                    <td><input type="text" value="${escapeHtml(t.cls || "")}" data-field="cls" data-idx="${idx}" placeholder="—"></td>
                    <td>
                        <div class="admin-actions">
                            <button class="icon-btn" data-action="up" data-idx="${idx}" title="Выше">▲</button>
                            <button class="icon-btn" data-action="down" data-idx="${idx}" title="Ниже">▼</button>
                            <button class="icon-btn" data-action="clear" data-idx="${idx}" title="Очистить">🧹</button>
                            <button class="icon-btn danger" data-action="delete" data-idx="${idx}" title="Удалить">🗑</button>
                        </div>
                    </td>
                `;
                tbody.appendChild(tr);
            });

            tbody.querySelectorAll("input").forEach(inp => {
                inp.addEventListener("change", (e) => {
                    const idx = parseInt(e.target.dataset.idx);
                    const field = e.target.dataset.field;
                    const oldVal = scheduleData[idx][field] || "";
                    const newVal = e.target.value.trim();
                    if (oldVal === newVal) return;
                    scheduleData[idx][field] = newVal;
                    addHistoryEntry({
                        type: "teacherMeta",
                        teacherIndex: idx,
                        teacherName: scheduleData[idx].name,
                        field, oldValue: oldVal, newValue: newVal
                    });
                    persistSchedule();
                    renderTable();
                    populateTeacherSelectForType();
                });
            });

            tbody.querySelectorAll(".icon-btn").forEach(btn => {
                btn.addEventListener("click", (e) => {
                    const idx = parseInt(e.target.dataset.idx);
                    handleAdminAction(e.target.dataset.action, idx);
                });
            });
        }

        function handleAdminAction(action, idx) {
            if (action === "up") {
                if (idx === 0) return;
                [scheduleData[idx - 1], scheduleData[idx]] = [scheduleData[idx], scheduleData[idx - 1]];
                persistSchedule();
                renderAdminTable(); renderTable(); populateTeacherSelectForType();
            } else if (action === "down") {
                if (idx === scheduleData.length - 1) return;
                [scheduleData[idx + 1], scheduleData[idx]] = [scheduleData[idx], scheduleData[idx + 1]];
                persistSchedule();
                renderAdminTable(); renderTable(); populateTeacherSelectForType();
            } else if (action === "clear") {
                const t = scheduleData[idx];
                if (!confirm(`Очистить всё расписание учителя «${t.name}»?`)) return;
                const snapshot = JSON.stringify(scheduleData);
                DAYS.forEach(day => {
                    t.lessons[day] = Array(MAX_LESSONS).fill(null).map(() => ({ text: "", substitute: "" }));
                });
                addHistoryEntry({ type: "bulk", description: `Очищено расписание: ${t.name}`, snapshot, affected: 1 });
                persistSchedule();
                renderTable();
                updateSubstituteIndicator();
                showToast(`Расписание «${t.name}» очищено`, "success");
            } else if (action === "delete") {
                const t = scheduleData[idx];
                if (!confirm(`Удалить учителя «${t.name}»?`)) return;
                const snapshot = JSON.parse(JSON.stringify(t));
                scheduleData.splice(idx, 1);
                addHistoryEntry({ type: "removeTeacher", teacherIndex: idx, teacherName: t.name, teacherSnapshot: snapshot });
                persistSchedule();
                renderAdminTable(); renderTable(); populateTeacherSelectForType();
                updateSubstituteIndicator();
                showToast(`Учитель «${t.name}» удалён`, "warning");
            }
        }

        function openTeacherEditModal(idx = -1) {
            editingTeacherIndex = idx;
            const modal = document.getElementById("teacherEditModal");
            const title = document.getElementById("teacherEditTitle");
            const desc = document.getElementById("teacherEditDesc");
            const nameInp = document.getElementById("teacherNameInput");
            const roomInp = document.getElementById("teacherRoomInput");
            const clsInp = document.getElementById("teacherClassInput");
            const err = document.getElementById("teacherEditError");

            err.textContent = "";
            if (idx === -1) {
                title.textContent = "Добавить учителя";
                desc.textContent = "Заполните поля и нажмите «Сохранить».";
                nameInp.value = ""; roomInp.value = ""; clsInp.value = "";
            } else {
                const t = scheduleData[idx];
                title.textContent = "Редактировать учителя";
                desc.textContent = "Измените поля и нажмите «Сохранить».";
                nameInp.value = t.name; roomInp.value = t.room || ""; clsInp.value = t.cls || "";
            }
            modal.classList.add("visible");
            setTimeout(() => nameInp.focus(), 50);
        }

        function closeTeacherEditModal() {
            document.getElementById("teacherEditModal").classList.remove("visible");
            editingTeacherIndex = -1;
        }

        function saveTeacherFromModal() {
            const name = document.getElementById("teacherNameInput").value.trim();
            const room = document.getElementById("teacherRoomInput").value.trim();
            const cls = document.getElementById("teacherClassInput").value.trim();
            const err = document.getElementById("teacherEditError");

            if (!name) { err.textContent = "❌ Имя учителя обязательно"; return; }

            if (editingTeacherIndex === -1) {
                const newTeacher = { name, room, cls, lessons: {} };
                DAYS.forEach(day => {
                    newTeacher.lessons[day] = Array(MAX_LESSONS).fill(null).map(() => ({ text: "", substitute: "" }));
                });
                scheduleData.push(newTeacher);
                addHistoryEntry({ type: "addTeacher", teacherIndex: scheduleData.length - 1, teacherName: name, room, cls });
                showToast(`✓ Учитель «${name}» добавлен`, "success");
            } else {
                const t = scheduleData[editingTeacherIndex];
                if (t.name !== name) {
                    addHistoryEntry({ type: "teacherMeta", teacherIndex: editingTeacherIndex, teacherName: name, field: "name", oldValue: t.name, newValue: name });
                    t.name = name;
                }
                if ((t.room || "") !== room) {
                    addHistoryEntry({ type: "teacherMeta", teacherIndex: editingTeacherIndex, teacherName: t.name, field: "room", oldValue: t.room || "", newValue: room });
                    t.room = room;
                }
                if ((t.cls || "") !== cls) {
                    addHistoryEntry({ type: "teacherMeta", teacherIndex: editingTeacherIndex, teacherName: t.name, field: "cls", oldValue: t.cls || "", newValue: cls });
                    t.cls = cls;
                }
                showToast("✓ Изменения сохранены", "success");
            }

            persistSchedule();
            closeTeacherEditModal();
            renderAdminTable(); renderTable(); populateTeacherSelectForType();
        }

        function initBulkActions() {
            const actionSel = document.getElementById("bulkAction");
            const teacherPicker = document.getElementById("bulkTeacherPicker");
            const dayPicker = document.getElementById("bulkDayPicker");
            const applyBtn = document.getElementById("applyBulkBtn");

            actionSel.addEventListener("change", () => {
                const v = actionSel.value;
                teacherPicker.style.display = v === "clearTeacher" ? "block" : "none";
                dayPicker.style.display = v === "clearDay" ? "block" : "none";
                applyBtn.disabled = !v;
            });

            applyBtn.addEventListener("click", () => {
                const v = actionSel.value;
                if (!v) return;
                const snapshot = JSON.stringify(scheduleData);
                let affected = 0;

                if (v === "clearTeacher") {
                    const teacherName = document.getElementById("bulkTeacherSelect").value;
                    if (!teacherName) return;
                    const t = scheduleData.find(x => x.name === teacherName);
                    if (!t) return;
                    if (!confirm(`Очистить расписание «${teacherName}»?`)) return;
                    DAYS.forEach(day => { t.lessons[day] = Array(MAX_LESSONS).fill(null).map(() => ({ text: "", substitute: "" })); });
                    affected = 1;
                    addHistoryEntry({ type: "bulk", description: `Очищено расписание: ${teacherName}`, snapshot, affected });
                    showToast(`✓ Очищено «${teacherName}»`, "success");
                } else if (v === "clearDay") {
                    const day = document.getElementById("bulkDaySelect").value;
                    if (!confirm(`Очистить день «${day}» у всех?`)) return;
                    scheduleData.forEach(t => { t.lessons[day] = Array(MAX_LESSONS).fill(null).map(() => ({ text: "", substitute: "" })); });
                    affected = scheduleData.length;
                    addHistoryEntry({ type: "bulk", description: `Очищен день: ${day}`, snapshot, affected });
                    showToast(`✓ Очищен день «${day}»`, "success");
                } else if (v === "clearSubstitutes") {
                    if (!confirm("Убрать все замены?")) return;
                    let removed = 0;
                    scheduleData.forEach(t => {
                        DAYS.forEach(day => {
                            (t.lessons[day] || []).forEach(l => {
                                if (l && l.substitute && l.substitute.trim()) { l.substitute = ""; removed++; }
                            });
                        });
                    });
                    affected = removed;
                    addHistoryEntry({ type: "substituteBulk", description: "Убраны все замены", snapshot, affected: removed });
                    showToast(`✓ Убрано замен: ${removed}`, "success");
                } else if (v === "clearAll") {
                    if (!confirm("Очистить ВСЁ расписание?")) return;
                    scheduleData.forEach(t => {
                        DAYS.forEach(day => { t.lessons[day] = Array(MAX_LESSONS).fill(null).map(() => ({ text: "", substitute: "" })); });
                    });
                    affected = scheduleData.length;
                    addHistoryEntry({ type: "bulk", description: "Полная очистка расписания", snapshot, affected });
                    showToast("✓ Всё расписание очищено", "warning");
                }

                persistSchedule();
                renderTable();
                updateBulkStats();
                updateSubstituteIndicator();
                actionSel.value = "";
                teacherPicker.style.display = "none";
                dayPicker.style.display = "none";
                applyBtn.disabled = true;
            });
        }

        function updateBulkStats() {
            const container = document.getElementById("bulkStats");
            if (!container) return;

            let totalLessons = 0, totalSubstitutes = 0;
            const perDay = {};
            DAYS.forEach(d => perDay[d] = 0);

            scheduleData.forEach(t => {
                DAYS.forEach(day => {
                    (t.lessons[day] || []).forEach(l => {
                        const lesson = normalizeLesson(l);
                        if (lesson.text && lesson.text.trim()) { totalLessons++; perDay[day]++; }
                        if (lesson.substitute && lesson.substitute.trim()) totalSubstitutes++;
                    });
                });
            });

            container.innerHTML = `
                <div>📚 Всего уроков: <strong>${totalLessons}</strong></div>
                <div>📌 Замен: <strong style="color:var(--substitute-text)">${totalSubstitutes}</strong></div>
                <div>👥 Учителей: <strong>${scheduleData.length}</strong></div>
                <div style="margin-top:6px; display:flex; flex-wrap:wrap; gap:10px;">
                    ${DAYS.map(d => `<span>${d}: <strong>${perDay[d]}</strong></span>`).join("")}
                </div>
            `;
        }

        function initImportExport() {
            document.getElementById("exportScheduleBtn").addEventListener("click", () => {
                const data = {
                    exportedAt: new Date().toISOString(),
                    school: "МОУ «Северная СОШ №2»",
                    version: 13,
                    schedule: scheduleData,
                    classes: classesData,
                    iup: iupData,
                    substitutions: substitutionsJournal,
                    substitutionsHistory: substitutionsHistory,
                    dictTeachers: dictTeachersManual,
                    dictSubjects: dictSubjectsManual,
                    constructor: constructorConfig
                };
                downloadJSON(data, `raspisanie_${formatDateForFile()}.json`);
                showToast("💾 Файл расписания скачан", "success");
            });

            document.getElementById("exportHistoryBtn").addEventListener("click", () => {
                const data = { exportedAt: new Date().toISOString(), history: historyData, substitutionsHistory: substitutionsHistory };
                downloadJSON(data, `istoriya_${formatDateForFile()}.json`);
                showToast("📜 Файл истории скачан", "success");
            });

            document.getElementById("importBtn").addEventListener("click", () => {
                document.getElementById("importFile").click();
            });

            document.getElementById("importFile").addEventListener("change", (e) => {
                const file = e.target.files[0];
                if (!file) return;
                document.getElementById("importFileName").textContent = "📂 " + file.name;
                const reader = new FileReader();
                reader.onload = (ev) => {
                    try {
                        const parsed = JSON.parse(ev.target.result);
                        const imported = parsed.schedule || parsed;

                        if (Array.isArray(imported) && imported.length) {
                            if (!confirm(`Импортировать учительское расписание (${imported.length})?`)) return;
                            const snapshot = JSON.stringify(scheduleData);
                            scheduleData = imported;
                            addHistoryEntry({ type: "bulk", description: `Импорт из JSON: ${file.name}`, snapshot, affected: imported.length });
                            persistSchedule();
                        }

                        if (parsed.classes && Array.isArray(parsed.classes)) { classesData = parsed.classes; persistClasses(); }
                        if (parsed.iup && Array.isArray(parsed.iup)) { iupData = parsed.iup; persistIUP(); }
                        if (parsed.substitutions && Array.isArray(parsed.substitutions)) { substitutionsJournal = parsed.substitutions; persistSubstitutions(); }
                        if (parsed.substitutionsHistory && Array.isArray(parsed.substitutionsHistory)) { substitutionsHistory = parsed.substitutionsHistory; persistSubstitutionsHistory(); }
                        if (parsed.dictTeachers && Array.isArray(parsed.dictTeachers)) { dictTeachersManual = parsed.dictTeachers; persistDictTeachers(); }
                        if (parsed.dictSubjects && Array.isArray(parsed.dictSubjects)) { dictSubjectsManual = parsed.dictSubjects; persistDictSubjects(); }
                        if (parsed.constructor && typeof parsed.constructor === "object") {
                            constructorConfig = deepMerge(JSON.parse(JSON.stringify(FACTORY_CONFIG)), parsed.constructor);
                            persistConstructorConfig();
                            if (typeof applyConstructorConfig === "function") applyConstructorConfig();
                        }

                        renderAdminTable(); renderTable(); populateTeacherSelectForType(); updateBulkStats();
                        updateSubstituteIndicator();
                        renderSubstitutionsPanel();
                        showToast(`✓ Импорт выполнен`, "success");
                        document.getElementById("importFileName").textContent = "";
                        e.target.value = "";
                    } catch (err) {
                        showToast(`❌ Ошибка: ${err.message}`, "error");
                    }
                };
                reader.readAsText(file, "utf-8");
            });

            document.getElementById("copyJsonBtn").addEventListener("click", async () => {
                const data = { schedule: scheduleData, classes: classesData, iup: iupData, substitutions: substitutionsJournal, dictTeachers: dictTeachersManual, dictSubjects: dictSubjectsManual, constructor: constructorConfig };
                const json = JSON.stringify(data, null, 2);
                try {
                    await navigator.clipboard.writeText(json);
                    showToast("📋 JSON скопирован", "success");
                } catch (e) { showToast("❌ Не удалось скопировать", "error"); }
            });
        }

        function downloadJSON(data, filename) {
            const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url; a.download = filename;
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        function formatDateForFile() {
            const d = new Date();
            const pad = n => String(n).padStart(2, "0");
            return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}-${pad(d.getMinutes())}`;
        }

        function exportToExcel() {
            try {
                const wb = XLSX.utils.book_new();

                const aoa = [];
                const header1 = ["Учитель", "Каб.", "Кл."];
                DAYS.forEach(day => { for (let i = 0; i < MAX_LESSONS; i++) header1.push(i === 0 ? day : ""); });
                aoa.push(header1);
                const header2 = ["", "", ""];
                DAYS.forEach(() => { for (let i = 1; i <= MAX_LESSONS; i++) header2.push(i); });
                aoa.push(header2);
                scheduleData.forEach(t => {
                    const row = [t.name, t.room || "", t.cls || ""];
                    DAYS.forEach(day => {
                        for (let i = 0; i < MAX_LESSONS; i++) {
                            const lesson = normalizeLesson(t.lessons[day]?.[i]);
                            let cellText = lesson.text || "";
                            if (lesson.substitute && lesson.substitute.trim()) cellText += ` (замена у ${lesson.substitute})`;
                            row.push(cellText);
                        }
                    });
                    aoa.push(row);
                });
                const ws = XLSX.utils.aoa_to_sheet(aoa);
                const merges = [];
                for (let d = 0; d < DAYS.length; d++) {
                    const startCol = 3 + d * MAX_LESSONS;
                    merges.push({ s: { r: 0, c: startCol }, e: { r: 0, c: startCol + MAX_LESSONS - 1 } });
                }
                ws["!merges"] = merges;
                const colWidths = [{ wch: 22 }, { wch: 8 }, { wch: 8 }];
                for (let i = 0; i < DAYS.length * MAX_LESSONS; i++) colWidths.push({ wch: 14 });
                ws["!cols"] = colWidths;
                XLSX.utils.book_append_sheet(wb, ws, "Учительское");

                const aoaC = [];
                const h1C = ["Класс", "Каб."];
                DAYS.forEach(day => { for (let i = 0; i < MAX_LESSONS; i++) h1C.push(i === 0 ? day : ""); });
                aoaC.push(h1C);
                const h2C = ["", ""];
                DAYS.forEach(() => { for (let i = 1; i <= MAX_LESSONS; i++) h2C.push(i); });
                aoaC.push(h2C);
                classesData.forEach(c => {
                    const row = [c.name, c.room || ""];
                    DAYS.forEach(day => {
                        for (let i = 0; i < MAX_LESSONS; i++) {
                            const lesson = normalizeLesson(c.lessons[day]?.[i]);
                            let cellText = lesson.text || "";
                            if (lesson.substitute && lesson.substitute.trim()) cellText += ` (замена: ${lesson.substitute})`;
                            row.push(cellText);
                        }
                    });
                    aoaC.push(row);
                });
                const wsC = XLSX.utils.aoa_to_sheet(aoaC);
                const mergesC = [];
                for (let d = 0; d < DAYS.length; d++) {
                    const startCol = 2 + d * MAX_LESSONS;
                    mergesC.push({ s: { r: 0, c: startCol }, e: { r: 0, c: startCol + MAX_LESSONS - 1 } });
                }
                wsC["!merges"] = mergesC;
                XLSX.utils.book_append_sheet(wb, wsC, "Детское");

                const aoaI = [];
                const h1I = ["Ученик", "Класс"];
                DAYS.forEach(day => { for (let i = 0; i < MAX_LESSONS; i++) h1I.push(i === 0 ? day : ""); });
                aoaI.push(h1I);
                const h2I = ["", ""];
                DAYS.forEach(() => { for (let i = 1; i <= MAX_LESSONS; i++) h2I.push(i); });
                aoaI.push(h2I);
                iupData.forEach(s => {
                    const row = [s.name, s.className || ""];
                    DAYS.forEach(day => {
                        for (let i = 0; i < MAX_LESSONS; i++) {
                            const lesson = normalizeLesson(s.lessons[day]?.[i]);
                            let cellText = lesson.text || "";
                            if (lesson.substitute && lesson.substitute.trim()) cellText += ` (замена: ${lesson.substitute})`;
                            row.push(cellText);
                        }
                    });
                    aoaI.push(row);
                });
                const wsI = XLSX.utils.aoa_to_sheet(aoaI);
                const mergesI = [];
                for (let d = 0; d < DAYS.length; d++) {
                    const startCol = 2 + d * MAX_LESSONS;
                    mergesI.push({ s: { r: 0, c: startCol }, e: { r: 0, c: startCol + MAX_LESSONS - 1 } });
                }
                wsI["!merges"] = mergesI;
                XLSX.utils.book_append_sheet(wb, wsI, "ИУП");

                XLSX.writeFile(wb, `raspisanie_${formatDateForFile()}.xlsx`);
                showToast("📤 Excel-файл скачан (3 листа)", "success");
            } catch (err) {
                console.error(err);
                showToast(`❌ Ошибка экспорта: ${err.message}`, "error");
            }
        }

        function exportHistoryToExcel() {
            try {
                const wb = XLSX.utils.book_new();
                const aoa = [["Время", "Тип", "Раздел", "Учитель/Ученик", "День", "Урок", "Было", "Стало", "Отменено"]];
                historyData.forEach(entry => {
                    const time = new Date(entry.timestamp).toLocaleString("ru-RU");
                    let type = "", section = "", person = "", day = "", lesson = "", oldVal = "", newVal = "";
                    if (entry.type === "edit") {
                        type = "Изменение урока";
                        section = entry.scheduleLabel || "Учительское";
                        person = entry.teacherName;
                        day = entry.day;
                        lesson = entry.lessonIndex + 1;
                        oldVal = entry.oldValue || "";
                        if (entry.oldSubstitute) oldVal += ` (замена: ${entry.oldSubstitute})`;
                        newVal = entry.newValue || "";
                        if (entry.newSubstitute) newVal += ` (замена: ${entry.newSubstitute})`;
                    } else if (entry.type === "passwordChange") {
                        type = "Смена пароля";
                        oldVal = entry.description || "";
                    } else if (entry.type === "bulk") {
                        type = "Массовая операция";
                        oldVal = entry.description || "";
                    }
                    aoa.push([time, type, section, person, day, lesson, oldVal, newVal, entry.undone ? "Да" : "Нет"]);
                });
                const ws = XLSX.utils.aoa_to_sheet(aoa);
                XLSX.utils.book_append_sheet(wb, ws, "История");
                XLSX.writeFile(wb, `istoriya_${formatDateForFile()}.xlsx`);
                showToast("📜 Excel-история скачана", "success");
            } catch (err) {
                console.error(err);
                showToast(`❌ Ошибка: ${err.message}`, "error");
            }
        }

        function initExcelImport() {
            const importBtn = document.getElementById("importExcelBtn");
            const fileInput = document.getElementById("importExcelFile");
            const optsBlock = document.getElementById("excelImportOptions");
            const cancelBtn = document.getElementById("cancelExcelImportBtn");

            importBtn.addEventListener("click", () => fileInput.click());

            fileInput.addEventListener("change", (e) => {
                const file = e.target.files[0];
                if (!file) return;
                const reader = new FileReader();
                reader.onload = (ev) => {
                    try {
                        const data = new Uint8Array(ev.target.result);
                        excelWorkbook = XLSX.read(data, { type: "array" });
                        optsBlock.style.display = "block";
                        showToast(`📂 Файл загружен: ${file.name}`, "success");
                    } catch (err) {
                        showToast(`❌ Ошибка: ${err.message}`, "error");
                    }
                };
                reader.readAsArrayBuffer(file);
                fileInput.value = "";
            });

            cancelBtn.addEventListener("click", () => {
                optsBlock.style.display = "none";
                excelWorkbook = null;
            });
        }

        function initSecurityTab() {
            document.querySelectorAll(".toggle-visibility").forEach(btn => {
                btn.addEventListener("click", () => {
                    const targetId = btn.dataset.target;
                    const inp = document.getElementById(targetId);
                    if (!inp) return;
                    if (inp.type === "password") { inp.type = "text"; btn.textContent = "🙈"; }
                    else { inp.type = "password"; btn.textContent = "👁"; }
                });
            });

            const newPassInput = document.getElementById("newPasswordInput");
            const strengthBlock = document.getElementById("passwordStrength");
            const strengthLabel = document.getElementById("passwordStrengthLabel");

            function evaluateStrength(pwd) {
                if (!pwd) return { level: "", label: "Введите пароль" };
                let score = 0;
                if (pwd.length >= 4) score++;
                if (pwd.length >= 8) score++;
                if (pwd.length >= 12) score++;
                if (/[a-z]/.test(pwd)) score++;
                if (/[A-Z]/.test(pwd)) score++;
                if (/\d/.test(pwd)) score++;
                if (/[^a-zA-Z0-9]/.test(pwd)) score++;
                if (score <= 2) return { level: "weak", label: "Слабый пароль" };
                if (score <= 4) return { level: "medium", label: "Средний пароль" };
                return { level: "strong", label: "Надёжный пароль" };
            }

            if (newPassInput) {
                newPassInput.addEventListener("input", () => {
                    const { level, label } = evaluateStrength(newPassInput.value);
                    strengthBlock.className = "password-strength" + (level ? " " + level : "");
                    strengthLabel.className = "password-strength-label" + (level ? " " + level : "");
                    strengthLabel.textContent = label;
                });
            }

            document.getElementById("changePasswordBtn").addEventListener("click", () => {
                const currentInp = document.getElementById("currentPasswordInput");
                const newInp = document.getElementById("newPasswordInput");
                const confirmInp = document.getElementById("confirmPasswordInput");
                const msgEl = document.getElementById("passwordChangeMsg");

                msgEl.className = "password-msg";
                msgEl.textContent = "";

                const currentPwd = currentInp.value;
                const newPwd = newInp.value.trim();
                const confirmPwd = confirmInp.value.trim();

                if (!currentPwd) { msgEl.className = "password-msg error"; msgEl.textContent = "❌ Введите текущий пароль"; return; }
                if (currentPwd !== getCurrentPassword()) { msgEl.className = "password-msg error"; msgEl.textContent = "❌ Неверный пароль"; currentInp.value = ""; return; }
                if (newPwd.length < 4) { msgEl.className = "password-msg error"; msgEl.textContent = "❌ Минимум 4 символа"; return; }
                if (newPwd === currentPwd) { msgEl.className = "password-msg error"; msgEl.textContent = "❌ Совпадает с текущим"; return; }
                if (newPwd !== confirmPwd) { msgEl.className = "password-msg error"; msgEl.textContent = "❌ Пароли не совпадают"; return; }

                if (savePassword(newPwd)) {
                    msgEl.className = "password-msg success";
                    msgEl.textContent = "✅ Пароль изменён!";
                    currentInp.value = "";
                    newInp.value = "";
                    confirmInp.value = "";
                    addHistoryEntry({ type: "passwordChange", description: "Пароль изменён" });
                    updateCurrentPasswordInfo();
                    showToast("🔐 Пароль изменён", "success");
                }
            });

            document.getElementById("resetPasswordBtn").addEventListener("click", () => {
                const currentInp = document.getElementById("currentPasswordInput");
                const msgEl = document.getElementById("passwordChangeMsg");
                if (!currentInp.value) { msgEl.className = "password-msg error"; msgEl.textContent = "❌ Введите текущий пароль"; return; }
                if (currentInp.value !== getCurrentPassword()) { msgEl.className = "password-msg error"; msgEl.textContent = "❌ Неверный пароль"; return; }
                if (!confirm("Сбросить пароль к sever2?")) return;
                if (resetPasswordToDefault()) {
                    msgEl.className = "password-msg success";
                    msgEl.textContent = "✅ Пароль сброшен";
                    currentInp.value = "";
                    document.getElementById("newPasswordInput").value = "";
                    document.getElementById("confirmPasswordInput").value = "";
                    updateCurrentPasswordInfo();
                    showToast("↺ Пароль сброшен к sever2", "warning");
                }
            });

            document.getElementById("forceResetPasswordBtn").addEventListener("click", () => {
                if (!confirm("⚠ Сбросить пароль к sever2 без ввода текущего?")) return;
                if (!confirm("Это снизит безопасность. Продолжить?")) return;
                if (resetPasswordToDefault()) {
                    addHistoryEntry({ type: "passwordChange", description: "Принудительный сброс пароля" });
                    updateCurrentPasswordInfo();
                    showToast("🔓 Пароль сброшен к sever2", "warning");
                    document.getElementById("passwordChangeMsg").textContent = "✅ Пароль сброшен.";
                    document.getElementById("passwordChangeMsg").className = "password-msg success";
                }
            });

            updateCurrentPasswordInfo();
        }

        function updateCurrentPasswordInfo() {
            const container = document.getElementById("currentPasswordInfo");
            if (!container) return;
            let isDefault = false;
            try { isDefault = !localStorage.getItem(PASSWORD_KEY); } catch (e) { isDefault = true; }
            const pwd = getCurrentPassword();
            const length = pwd.length;
            const isDefaultPwd = pwd === DEFAULT_PASSWORD;
            let lengthText = "";
            if (length < 6) lengthText = `⚠ короткий (${length})`;
            else if (length < 10) lengthText = `средний (${length})`;
            else lengthText = `длинный (${length})`;
            container.innerHTML = `
                <div>📌 <strong>Тип:</strong> ${isDefaultPwd ? "стандартный (sever2)" : "пользовательский"}</div>
                <div>📏 <strong>Длина:</strong> ${lengthText} символов</div>
                <div>💾 <strong>Хранение:</strong> ${isDefault ? "в коде" : "в браузере"}</div>
            `;
        }

        function initDangerZone() {
            document.getElementById("fullResetBtn").addEventListener("click", () => {
                if (!confirm("⚠ Полный сброс ВСЕХ трёх расписаний к исходным данным?")) return;
                if (!confirm("Вы уверены?")) return;
                scheduleData = buildScheduleFromRaw();
                classesData = rawClasses.map(c => ({
                    name: c.name, room: c.room || "",
                    lessons: (() => {
                        const lessons = {};
                        DAYS.forEach(day => {
                            const arr = c.lessons[day] || [];
                            const padded = [];
                            for (let i = 0; i < MAX_LESSONS; i++) padded.push({ text: arr[i] || "", substitute: "" });
                            lessons[day] = padded;
                        });
                        return lessons;
                    })()
                }));
                iupData = rawIUP.map(s => ({
                    name: s.name, className: s.className || "",
                    lessons: (() => {
                        const lessons = {};
                        DAYS.forEach(day => {
                            const arr = s.lessons[day] || [];
                            const padded = [];
                            for (let i = 0; i < MAX_LESSONS; i++) padded.push({ text: arr[i] || "", substitute: "" });
                            lessons[day] = padded;
                        });
                        return lessons;
                    })()
                }));
                persistSchedule();
                persistClasses();
                persistIUP();
                addHistoryEntry({ type: "bulk", description: "Полный сброс всех расписаний" });
                renderAdminTable(); renderTable(); populateTeacherSelectForType(); updateBulkStats();
                updateSubstituteIndicator();
                showToast("↺ Все расписания сброшены", "warning");
            });

            document.getElementById("clearAllLocalBtn").addEventListener("click", () => {
                if (!confirm("⚠ Очистить ВСЁ локальное хранилище?")) return;
                if (!confirm("Это необратимо. Продолжить?")) return;
                Object.values(STORAGE_KEYS).forEach(k => localStorage.removeItem(k));
                localStorage.removeItem(PASSWORD_KEY);
                location.reload();
            });
        }

        function updateStorageInfo() {
            const container = document.getElementById("storageInfo");
            if (!container) return;
            try {
                const schedSize = (localStorage.getItem(STORAGE_KEYS.schedule) || "").length;
                const classesSize = (localStorage.getItem(STORAGE_KEYS.classes) || "").length;
                const iupSize = (localStorage.getItem(STORAGE_KEYS.iup) || "").length;
                const histSize = (localStorage.getItem(STORAGE_KEYS.history) || "").length;
                const substSize = (localStorage.getItem(STORAGE_KEYS.substitutions) || "").length;
                const substHistSize = (localStorage.getItem(STORAGE_KEYS.substitutionsHistory) || "").length;
                const dictTSize = (localStorage.getItem(STORAGE_KEYS.dictTeachers) || "").length;
                const dictSSize = (localStorage.getItem(STORAGE_KEYS.dictSubjects) || "").length;
                const constrSize = (localStorage.getItem(STORAGE_KEYS.constructor) || "").length;
                const updBackupSize = (localStorage.getItem(STORAGE_KEYS.updateBackup) || "").length;
                const updDataBackupSize = (localStorage.getItem(STORAGE_KEYS.updateDataBackup) || "").length;
                const totalSize = schedSize + classesSize + iupSize + histSize + substSize + substHistSize + dictTSize + dictSSize + constrSize + updBackupSize + updDataBackupSize;
                const fmt = b => b < 1024 ? `${b} Б` : `${(b / 1024).toFixed(1)} КБ`;
                container.innerHTML = `
                    <div>👨‍🏫 Учительское: <strong>${fmt(schedSize)}</strong></div>
                    <div>🎓 Детское: <strong>${fmt(classesSize)}</strong></div>
                    <div>📚 ИУП: <strong>${fmt(iupSize)}</strong></div>
                    <div>📜 История: <strong>${fmt(histSize)}</strong></div>
                    <div>📌 Журнал замен: <strong>${fmt(substSize)}</strong></div>
                    <div>📜 История замен: <strong>${fmt(substHistSize)}</strong></div>
                    <div>👥 Справочник учителей: <strong>${fmt(dictTSize)}</strong></div>
                    <div>📚 Справочник предметов: <strong>${fmt(dictSSize)}</strong></div>
                    <div>🎨 Конструктор: <strong>${fmt(constrSize)}</strong></div>
                    <div>🔄 Бэкап HTML: <strong>${fmt(updBackupSize)}</strong></div>
                    <div>🛡️ Бэкап данных: <strong>${fmt(updDataBackupSize)}</strong></div>
                    <div>💾 Всего: <strong>${fmt(totalSize)}</strong></div>
                `;
            } catch (e) { container.textContent = "Ошибка"; }
        }

        function initHistoryPanel() {
            const panel = document.getElementById("historyPanel");
            const toggleBtn = document.getElementById("toggleHistoryBtn");

            toggleBtn.addEventListener("click", () => {
                historyPanelExpanded = !historyPanelExpanded;
                panel.classList.toggle("visible", historyPanelExpanded);
                toggleBtn.textContent = historyPanelExpanded ? "Развернуть" : "Свернуть";
            });

            document.getElementById("showHistoryBtn").addEventListener("click", () => {
                historyPanelExpanded = true;
                panel.classList.add("visible");
                toggleBtn.textContent = "Свернуть";
            });

            document.getElementById("undoLastBtn").addEventListener("click", undoLastChange);
        }

        function initAnimatedBg() {
            const bg = document.getElementById("animatedBg");
            const words = [
                "Математика", "Русский язык", "Литература", "Физика", "Химия", "Биология",
                "История", "География", "Английский язык", "Информатика", "Физкультура",
                "Технология", "Музыка", "ИЗО", "ОБЗР", "Обществознание", "Алгебра", "Геометрия",
                "Вероятность и статистика", "Алгоритмика", "Проект", "Разговоры о важном"
            ];
            const days = ["Пн", "Вт", "Ср", "Чт", "Пт"];
            const numbers = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11"];

            for (let i = 0; i < 22; i++) {
                const el = document.createElement("div");
                el.className = "word";
                el.textContent = words[Math.floor(Math.random() * words.length)];
                el.style.left = Math.random() * 100 + "%";
                el.style.fontSize = (Math.random() * 2.5 + 1.2) + "rem";
                el.style.animationDuration = (Math.random() * 30 + 25) + "s";
                el.style.animationDelay = (Math.random() * 25) + "s";
                bg.appendChild(el);
            }
            for (let i = 0; i < 12; i++) {
                const el = document.createElement("div");
                el.className = "number";
                el.textContent = numbers[Math.floor(Math.random() * numbers.length)];
                el.style.left = Math.random() * 100 + "%";
                el.style.fontSize = (Math.random() * 4 + 2) + "rem";
                el.style.animationDuration = (Math.random() * 35 + 30) + "s";
                el.style.animationDelay = (Math.random() * 30) + "s";
                bg.appendChild(el);
            }
            for (let i = 0; i < 8; i++) {
                const el = document.createElement("div");
                el.className = "day";
                el.textContent = days[Math.floor(Math.random() * days.length)];
                el.style.left = Math.random() * 100 + "%";
                el.style.fontSize = (Math.random() * 3 + 2) + "rem";
                el.style.animationDuration = (Math.random() * 40 + 30) + "s";
                el.style.animationDelay = (Math.random() * 30) + "s";
                bg.appendChild(el);
            }
        }

        const DEFAULT_BELLS = [
            { start: "08:30", end: "09:15" },
            { start: "09:25", end: "10:10" },
            { start: "10:25", end: "11:10" },
            { start: "11:25", end: "12:10" },
            { start: "12:25", end: "13:10" },
            { start: "13:20", end: "14:05" },
            { start: "14:15", end: "15:00" },
            { start: "15:10", end: "15:55" }
        ];

        function loadBells() {
            try {
                const saved = localStorage.getItem(STORAGE_KEYS.bells);
                if (saved) {
                    const arr = JSON.parse(saved);
                    if (Array.isArray(arr) && arr.length === MAX_LESSONS) return arr;
                }
            } catch (e) {}
            return JSON.parse(JSON.stringify(DEFAULT_BELLS));
        }

        function saveBells(bells) {
            localStorage.setItem(STORAGE_KEYS.bells, JSON.stringify(bells));
        }

        let currentBells = loadBells();

        function getTodayDayName() {
            const dayIdx = new Date().getDay();
            if (dayIdx === 0 || dayIdx === 6) return null;
            return DAYS[dayIdx - 1];
        }

        function getTomorrowDayName() {
            const dayIdx = (new Date().getDay() + 1) % 7;
            if (dayIdx === 0 || dayIdx === 6) return null;
            return DAYS[dayIdx - 1];
        }

        function getMyTeacherName() {
            return localStorage.getItem(STORAGE_KEYS.myTeacher) || "";
        }

        function openMySchedule() {
            const modal = document.getElementById("myScheduleModal");
            const titleEl = modal.querySelector("h2");
            const teacherControl = document.querySelector(".my-sched-control:nth-child(1)");
            const labelEl = teacherControl.querySelector("label");
            const selectEl = document.getElementById("myScheduleTeacher");

            if (currentScheduleType === "teachers" || currentScheduleType === "substitutions") {
                if (titleEl) titleEl.textContent = "👤 Моё расписание (учитель)";
                if (labelEl) labelEl.textContent = "Учитель:";
                selectEl.innerHTML = '<option value="">— Выберите учителя —</option>';
                const names = scheduleData.map(t => t.name).sort((a, b) => a.localeCompare(b, "ru"));
                names.forEach(name => {
                    const opt = document.createElement("option");
                    opt.value = name;
                    opt.textContent = name;
                    selectEl.appendChild(opt);
                });
                const currentVal = getMyTeacherName();
                if (currentVal && names.includes(currentVal)) selectEl.value = currentVal;
                else selectEl.value = "";
            } else if (currentScheduleType === "classes") {
                if (titleEl) titleEl.textContent = "🎓 Расписание класса";
                if (labelEl) labelEl.textContent = "Класс:";
                selectEl.innerHTML = '<option value="">— Выберите класс —</option>';
                const names = classesData.map(c => c.name).sort((a, b) => a.localeCompare(b, "ru", { numeric: true }));
                names.forEach(name => {
                    const opt = document.createElement("option");
                    opt.value = name;
                    opt.textContent = name;
                    selectEl.appendChild(opt);
                });
                const savedClass = localStorage.getItem(STORAGE_KEYS.myClass) || "";
                if (savedClass && names.includes(savedClass)) selectEl.value = savedClass;
                else selectEl.value = "";
            } else if (currentScheduleType === "iup") {
                if (titleEl) titleEl.textContent = "📚 Индивидуальный план ученика";
                if (labelEl) labelEl.textContent = "Ученик:";
                selectEl.innerHTML = '<option value="">— Выберите ученика —</option>';
                const names = iupData.map(s => s.name).sort((a, b) => a.localeCompare(b, "ru"));
                names.forEach(name => {
                    const opt = document.createElement("option");
                    opt.value = name;
                    opt.textContent = name;
                    selectEl.appendChild(opt);
                });
                const savedStudent = localStorage.getItem(STORAGE_KEYS.myStudent) || "";
                if (savedStudent && names.includes(savedStudent)) selectEl.value = savedStudent;
                else selectEl.value = "";
            }

            modal.classList.add("visible");
            renderMySchedule();
        }

        function closeMySchedule() {
            document.getElementById("myScheduleModal").classList.remove("visible");
        }

        function renderMySchedule() {
            const selectedName = document.getElementById("myScheduleTeacher").value;
            const dayValue = document.getElementById("myScheduleDay").value;
            const content = document.getElementById("myScheduleContent");

            if (!selectedName) {
                let hint = "👆 Выберите учителя";
                if (currentScheduleType === "classes") hint = "👆 Выберите класс";
                if (currentScheduleType === "iup") hint = "👆 Выберите ученика";
                content.innerHTML = `<div class="my-sched-hint">${hint}</div>`;
                return;
            }

            let dataArray = scheduleData;
            let savedKey = STORAGE_KEYS.myTeacher;
            if (currentScheduleType === "classes") {
                dataArray = classesData;
                savedKey = STORAGE_KEYS.myClass;
            } else if (currentScheduleType === "iup") {
                dataArray = iupData;
                savedKey = STORAGE_KEYS.myStudent;
            }

            const item = dataArray.find(t => t.name === selectedName);
            if (!item) {
                content.innerHTML = '<div class="my-sched-empty">Не найдено</div>';
                return;
            }

            localStorage.setItem(savedKey, selectedName);

            let daysToShow = [];
            if (dayValue === "today") {
                const today = getTodayDayName();
                if (today) daysToShow = [today];
                else { content.innerHTML = '<div class="my-sched-empty">Сегодня выходной 🌿</div>'; return; }
            } else if (dayValue === "tomorrow") {
                const tomorrow = getTomorrowDayName();
                if (tomorrow) daysToShow = [tomorrow];
                else { content.innerHTML = '<div class="my-sched-empty">Завтра выходной 🌿</div>'; return; }
            } else if (dayValue === "week") daysToShow = DAYS;
            else daysToShow = [dayValue];

            let html = "";
            daysToShow.forEach(day => {
                html += renderMyScheduleDay(item, day);
            });

            if (currentScheduleType === "teachers" || currentScheduleType === "substitutions") {
                const replacementsHtml = renderReplacementsForMe(item, daysToShow);
                if (replacementsHtml) html += replacementsHtml;
            }

            content.innerHTML = html || '<div class="my-sched-empty">Нет уроков</div>';
        }

        function renderMyScheduleDay(item, day) {
            const lessons = item.lessons[day] || [];
            let hasAnyLesson = false;
            let substitutionsCount = 0;
            let rowsHtml = "";

            for (let i = 0; i < MAX_LESSONS; i++) {
                const lesson = normalizeLesson(lessons[i]);
                const hasText = lesson.text && lesson.text.trim();
                const hasSubst = lesson.substitute && lesson.substitute.trim();
                const bell = currentBells[i] || { start: "", end: "" };
                const timeStr = bell.start && bell.end ? `${bell.start}–${bell.end}` : "—";

                if (!hasText && !hasSubst) {
                    rowsHtml += `<div class="my-sched-lesson empty"><div class="num">${i + 1}</div><div class="time">${timeStr}</div><div class="subject">— нет урока —</div><div class="class-name"></div><div class="room"></div></div>`;
                    continue;
                }

                hasAnyLesson = true;
                let subject = lesson.text || "";

                let substNoteText = "";
                if (hasSubst) {
                    substitutionsCount++;
                    if (currentScheduleType === "teachers" || currentScheduleType === "substitutions") {
                        substNoteText = `(замена у ${lesson.substitute})`;
                    } else {
                        substNoteText = `(замена: ${lesson.substitute})`;
                    }
                    rowsHtml += `<div class="my-sched-lesson has-substitute"><div class="num">${i + 1}</div><div class="time">${timeStr}</div><div class="subject"><span class="subst-icon">📌</span>${escapeHtml(subject)}<span class="subst-note">${escapeHtml(substNoteText)}</span></div><div class="class-name">—</div><div class="room"></div></div>`;
                } else {
                    rowsHtml += `<div class="my-sched-lesson"><div class="num">${i + 1}</div><div class="time">${timeStr}</div><div class="subject">${escapeHtml(subject)}</div><div class="class-name">—</div><div class="room"></div></div>`;
                }
            }

            if (!hasAnyLesson) {
                return `<div class="my-sched-day"><div class="my-sched-day-header"><span>${day}</span><span style="font-size:0.75rem; opacity:0.8;">Свободный день</span></div><div class="my-sched-lessons"><div class="my-sched-hint">🌸 Уроков нет</div></div></div>`;
            }

            return `<div class="my-sched-day"><div class="my-sched-day-header"><span>${day}</span>${substitutionsCount > 0 ? `<span class="substitutions-count">📌 Замен: ${substitutionsCount}</span>` : ""}</div><div class="my-sched-lessons">${rowsHtml}</div></div>`;
        }

        function renderReplacementsForMe(teacher, daysToShow) {
            const myName = teacher.name;
            let rows = [];
            daysToShow.forEach(day => {
                scheduleData.forEach(otherTeacher => {
                    if (otherTeacher.name === myName) return;
                    const lessons = otherTeacher.lessons[day] || [];
                    for (let i = 0; i < MAX_LESSONS; i++) {
                        const lesson = normalizeLesson(lessons[i]);
                        if (lesson.substitute && lesson.substitute === myName) {
                            rows.push({ day, lessonIdx: i, subject: lesson.text || "", otherTeacher: otherTeacher.name });
                        }
                    }
                });
            });

            if (rows.length === 0) return "";

            let html = `<div class="my-sched-replacements"><h3>🔔 В эти уроки вас заменяют</h3>`;
            rows.forEach(r => {
                const bell = currentBells[r.lessonIdx] || { start: "", end: "" };
                const timeStr = bell.start && bell.end ? `${bell.start}–${bell.end}` : "—";
                const dayShort = DAY_SHORT[r.day] || r.day;
                html += `<div class="my-sched-replacement-row"><div class="num">${r.lessonIdx + 1}</div><div class="time">${dayShort} · ${timeStr}</div><div><strong>${escapeHtml(r.subject)}</strong></div><div>ведёт <span class="who">${escapeHtml(r.otherTeacher)}</span></div></div>`;
            });
            html += "</div>";
            return html;
        }

        function openBellsModal() {
            const editor = document.getElementById("bellsEditor");
            editor.innerHTML = "";
            for (let i = 0; i < MAX_LESSONS; i++) {
                const bell = currentBells[i] || { start: "", end: "" };
                editor.innerHTML += `<div style="font-weight:700; color:var(--accent); text-align:center;">${i + 1}.</div><input type="time" data-idx="${i}" data-field="start" value="${bell.start || ""}" style="padding:6px 8px; border-radius:6px; border:1.5px solid var(--border); background:var(--table-bg); color:var(--text-primary); font-family:inherit;"><input type="time" data-idx="${i}" data-field="end" value="${bell.end || ""}" style="padding:6px 8px; border-radius:6px; border:1.5px solid var(--border); background:var(--table-bg); color:var(--text-primary); font-family:inherit;">`;
            }
            document.getElementById("bellsModal").classList.add("visible");
        }

        function closeBellsModal() {
            document.getElementById("bellsModal").classList.remove("visible");
        }

        function saveBellsFromModal() {
            const editor = document.getElementById("bellsEditor");
            const inputs = editor.querySelectorAll("input");
            const newBells = JSON.parse(JSON.stringify(DEFAULT_BELLS));
            inputs.forEach(inp => {
                const idx = parseInt(inp.dataset.idx);
                const field = inp.dataset.field;
                newBells[idx][field] = inp.value || "";
            });
            currentBells = newBells;
            saveBells(currentBells);
            closeBellsModal();
            renderMySchedule();
            showToast("🔔 Расписание звонков сохранено", "success");
        }

        function initMySchedule() {
            document.getElementById("myScheduleBtn").addEventListener("click", openMySchedule);
            document.getElementById("closeMyScheduleBtn").addEventListener("click", closeMySchedule);
            document.getElementById("myScheduleModal").addEventListener("click", (e) => {
                if (e.target.id === "myScheduleModal") closeMySchedule();
            });
            document.getElementById("myScheduleTeacher").addEventListener("change", renderMySchedule);
            document.getElementById("myScheduleDay").addEventListener("change", renderMySchedule);
            document.getElementById("printMyScheduleBtn").addEventListener("click", () => window.print());
            document.getElementById("bellsSettingsBtn").addEventListener("click", openBellsModal);
            document.getElementById("cancelBellsBtn").addEventListener("click", closeBellsModal);
            document.getElementById("saveBellsBtn").addEventListener("click", saveBellsFromModal);
            document.getElementById("bellsModal").addEventListener("click", (e) => {
                if (e.target.id === "bellsModal") closeBellsModal();
            });
        }

        // ==================== ЖУРНАЛ ЗАМЕН: ЛОГИКА ====================

        function getDayNameFromDate(d) {
            const idx = d.getDay();
            if (idx === 0 || idx === 6) return "ВЫХОДНОЙ";
            return DAYS[idx - 1];
        }

        function addSubstHistoryEntry(entry) {
            const record = {
                id: "sh_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8),
                timestamp: new Date().toISOString(),
                undone: false,
                ...entry
            };
            substitutionsHistory.unshift(record);
            if (substitutionsHistory.length > 500) substitutionsHistory = substitutionsHistory.slice(0, 500);
            persistSubstitutionsHistory();
            renderSubstHistory();
        }

        function undoSubstHistoryEntry(id) {
            const idx = substitutionsHistory.findIndex(h => h.id === id);
            if (idx === -1) return;
            const entry = substitutionsHistory[idx];
            if (entry.undone) { showToast("Это изменение уже отменено", "warning"); return; }

            if (entry.type === "add") {
                substitutionsJournal = substitutionsJournal.filter(r => r.id !== entry.recordId);
            } else if (entry.type === "delete") {
                if (entry.snapshot) substitutionsJournal.splice(entry.index, 0, entry.snapshot);
            } else if (entry.type === "edit") {
                const rec = substitutionsJournal.find(r => r.id === entry.recordId);
                if (rec) Object.assign(rec, entry.snapshotBefore);
            }

            entry.undone = true;
            persistSubstitutions();
            persistSubstitutionsHistory();
            renderSubstitutionsPanel();
            showToast("✓ Изменение журнала отменено", "success");
        }

        function undoLastSubstChange() {
            const last = substitutionsHistory.find(h => !h.undone);
            if (!last) { showToast("Нет изменений для отмены", "warning"); return; }
            undoSubstHistoryEntry(last.id);
        }

        function renderSubstHistory() {
            const container = document.getElementById("substHistoryList");
            const countEl = document.getElementById("substHistoryCount");
            if (!container) return;

            const active = substitutionsHistory.filter(h => !h.undone).length;
            if (countEl) countEl.textContent = active;

            if (substitutionsHistory.length === 0) {
                container.innerHTML = '<div class="history-empty">📭 История журнала замен пуста</div>';
                return;
            }

            container.innerHTML = "";
            substitutionsHistory.slice(0, 100).forEach(entry => {
                const item = document.createElement("div");
                item.className = "history-item" + (entry.undone ? " undone" : "");
                item.style.background = "var(--substitute-bg)";
                item.style.borderColor = "var(--substitute-border)";

                const time = document.createElement("div");
                time.className = "history-time";
                time.textContent = formatTime(entry.timestamp);

                const content = document.createElement("div");
                content.className = "history-content";

                let title = "", detail = "";
                if (entry.type === "add") {
                    title = `➕ Добавлена запись`;
                    detail = escapeHtml(entry.description || "");
                } else if (entry.type === "edit") {
                    title = `✎ Изменена запись`;
                    detail = escapeHtml(entry.description || "");
                } else if (entry.type === "delete") {
                    title = `🗑 Удалена запись`;
                    detail = escapeHtml(entry.description || "");
                } else {
                    title = escapeHtml(entry.description || entry.type);
                }

                content.innerHTML = `<div>${title}</div><div class="detail">${detail}</div>`;

                const actions = document.createElement("div");
                if (!entry.undone) {
                    const undoBtn = document.createElement("button");
                    undoBtn.className = "icon-btn";
                    undoBtn.title = "Откатить";
                    undoBtn.textContent = "↶";
                    undoBtn.addEventListener("click", () => undoSubstHistoryEntry(entry.id));
                    actions.appendChild(undoBtn);
                } else {
                    actions.innerHTML = '<span style="color:var(--text-muted); font-size:0.7rem;">отменено</span>';
                }

                item.appendChild(time);
                item.appendChild(content);
                item.appendChild(actions);
                container.appendChild(item);
            });
        }

        function getAllTeachersForJournal() {
            const set = new Set();
            scheduleData.forEach(t => { if (t.name) set.add(t.name.trim()); });
            dictTeachersManual.forEach(n => { if (n) set.add(n.trim()); });
            substitutionsJournal.forEach(r => {
                if (r.absentTeacher) set.add(r.absentTeacher.trim());
                if (r.substituteTeacher) set.add(r.substituteTeacher.trim());
            });
            return [...set].sort((a, b) => a.localeCompare(b, "ru"));
        }

        function getAllSubjectsForJournal() {
            const set = new Set();
            classesData.forEach(c => {
                DAYS.forEach(day => {
                    (c.lessons[day] || []).forEach(l => {
                        const lesson = normalizeLesson(l);
                        const txt = (lesson.text || "").trim();
                        if (txt) set.add(txt);
                    });
                });
            });
            dictSubjectsManual.forEach(n => { if (n) set.add(n.trim()); });
            substitutionsJournal.forEach(r => {
                if (r.absentSubject) set.add(r.absentSubject.trim());
                if (r.substituteSubject) set.add(r.substituteSubject.trim());
            });
            return [...set].sort((a, b) => a.localeCompare(b, "ru"));
        }

        function getSubstFiltered() {
            let list = substitutionsJournal.slice();

            if (substViewMode === "day" && currentSubstDay) {
                list = list.filter(r => (r.date || "") === currentSubstDay);
            } else if (substViewMode === "month" && currentSubstMonth) {
                list = list.filter(r => (r.date || "").startsWith(currentSubstMonth));
            }

            if (currentSubstSearch) {
                const q = currentSubstSearch.toLowerCase();
                list = list.filter(r =>
                    (r.absentTeacher || "").toLowerCase().includes(q) ||
                    (r.substituteTeacher || "").toLowerCase().includes(q) ||
                    (r.absentSubject || "").toLowerCase().includes(q) ||
                    (r.substituteSubject || "").toLowerCase().includes(q) ||
                    (r.lessonType || "").toLowerCase().includes(q)
                );
            }
            return list;
        }

        function populateSubstFormSelects() {
            const teachers = getAllTeachersForJournal();
            const subjects = getAllSubjectsForJournal();

            const fill = (selectId, placeholder, list) => {
                const sel = document.getElementById(selectId);
                if (!sel) return;
                const current = sel.value;
                sel.innerHTML = `<option value="">${placeholder}</option>`;
                list.forEach(v => {
                    const o = document.createElement("option");
                    o.value = v; o.textContent = v;
                    sel.appendChild(o);
                });
                if (current && [...sel.options].some(o => o.value === current)) sel.value = current;
            };

            fill("substAbsentTeacher", "— Кого заменяем —", teachers);
            fill("substSubstituteTeacher", "— Кто заменяет —", teachers);
            fill("substAbsentSubject", "— Предмет отсутствующего —", subjects);
            fill("substSubstituteSubject", "— Предмет заменяющего —", subjects);
        }

        function updateSubstDateHint() {
            const hintEl = document.getElementById("substDateHint");
            const dateInp = document.getElementById("substDateInput");
            if (!hintEl || !dateInp) return;

            if (substViewMode === "day" && currentSubstDay) {
                hintEl.textContent = "↻ Дата автоматически подставлена из выбранного дня";
                hintEl.classList.add("active");
                if (!dateInp.value || dateInp.value !== currentSubstDay) {
                    dateInp.value = currentSubstDay;
                }
            } else {
                hintEl.textContent = "";
                hintEl.classList.remove("active");
            }
        }

        function renderSubstitutionsPanel() {
            if (currentScheduleType !== "substitutions") return;

            populateSubstFormSelects();
            updateSubstDateHint();

            const filtered = getSubstFiltered();

            const sectionTitleEl = document.getElementById("substSectionTitle");
            if (sectionTitleEl) {
                if (substViewMode === "day" && currentSubstDay) {
                    const d = new Date(currentSubstDay + "T00:00:00");
                    const dayName = getDayNameFromDate(d);
                    const dayShort = DAY_SHORT[dayName] || dayName;
                    const dateStr = d.toLocaleDateString("ru-RU", { day: "numeric", month: "long", year: "numeric" });
                    sectionTitleEl.textContent = `📋 Журнал замен · ${dayShort}, ${dateStr}`;
                } else if (substViewMode === "month" && currentSubstMonth) {
                    const [y, m] = currentSubstMonth.split("-");
                    const monthName = new Date(parseInt(y), parseInt(m) - 1, 1)
                        .toLocaleDateString("ru-RU", { month: "long", year: "numeric" });
                    sectionTitleEl.textContent = `📋 Журнал замен · ${monthName}`;
                } else {
                    sectionTitleEl.textContent = "📋 Журнал замен";
                }
            }

            const totalEl = document.getElementById("substStatTotal");
            const replacedEl = document.getElementById("substStatReplaced");
            const subsEl = document.getElementById("substStatSubstitutes");
            const monthEl = document.getElementById("substStatMonth");

            const uniqueReplaced = new Set();
            const uniqueSubstitutes = new Set();
            const nowMonth = new Date().toISOString().slice(0, 7);
            let monthCount = 0;

            filtered.forEach(r => {
                if (r.absentTeacher) uniqueReplaced.add(r.absentTeacher);
                if (r.substituteTeacher) uniqueSubstitutes.add(r.substituteTeacher);
                if ((r.date || "").startsWith(nowMonth)) monthCount++;
            });

            if (totalEl) totalEl.textContent = filtered.length;
            if (replacedEl) replacedEl.textContent = uniqueReplaced.size;
            if (subsEl) subsEl.textContent = uniqueSubstitutes.size;
            if (monthEl) monthEl.textContent = monthCount;

            const tbody = document.getElementById("substDetailsBody");
            const countEl = document.getElementById("substDetailsCount");
            if (countEl) countEl.textContent = filtered.length;

            const details = filtered.slice().sort((a, b) => {
                if ((a.date || "") !== (b.date || "")) return (b.date || "").localeCompare(a.date || "");
                return (b.createdAt || "").localeCompare(a.createdAt || "");
            });

            tbody.innerHTML = "";
            if (details.length === 0) {
                let emptyMsg = "Записей нет. Добавьте первую замену выше. 👆";
                if (substViewMode === "day" && currentSubstDay) {
                    const d = new Date(currentSubstDay + "T00:00:00");
                    emptyMsg = `На ${d.toLocaleDateString("ru-RU", { day: "numeric", month: "long", year: "numeric" })} записей нет.`;
                }
                tbody.innerHTML = `<tr class="empty-row"><td colspan="8">${emptyMsg}</td></tr>`;
            } else {
                details.forEach(r => {
                    const tr = document.createElement("tr");
                    const dayName = r.day || (r.date ? getDayNameFromDate(new Date(r.date + "T00:00:00")) : "");
                    const dayShort = DAY_SHORT[dayName] || dayName || "—";
                    const dateStr = r.date ? new Date(r.date).toLocaleDateString("ru-RU") : "—";
                    const lessonType = r.lessonType || "Стандартный урок";
                    const typeBadgeClass = lessonType === "Внеурочная деятельность" ? "subst-type-badge extra" : "subst-type-badge standard";
                    tr.innerHTML = `
                        <td>${escapeHtml(dateStr)}</td>
                        <td style="text-align:center;">${escapeHtml(dayShort)}</td>
                        <td><span class="${typeBadgeClass}">${escapeHtml(lessonType)}</span></td>
                        <td><strong>${escapeHtml(r.absentTeacher || "—")}</strong></td>
                        <td>${escapeHtml(r.absentSubject || "—")}</td>
                        <td><strong>${escapeHtml(r.substituteTeacher || "—")}</strong></td>
                        <td>${escapeHtml(r.substituteSubject || "—")}</td>
                        <td class="actions">
                            <button class="icon-btn" data-action="edit" data-id="${r.id}" title="Редактировать">✎</button>
                            <button class="icon-btn danger" data-action="delete" data-id="${r.id}" title="Удалить">🗑</button>
                        </td>
                    `;
                    tbody.appendChild(tr);
                });

                tbody.querySelectorAll("[data-action='edit']").forEach(el => {
                    el.addEventListener("click", () => startEditSubstRecord(el.dataset.id));
                });
                tbody.querySelectorAll("[data-action='delete']").forEach(el => {
                    el.addEventListener("click", () => deleteSubstRecord(el.dataset.id));
                });
            }

            const fc = document.getElementById("substFilterCount");
            if (fc) {
                if (substViewMode === "day" && currentSubstDay) {
                    const d = new Date(currentSubstDay + "T00:00:00");
                    fc.textContent = `Показано: ${filtered.length} за ${d.toLocaleDateString("ru-RU")}`;
                } else {
                    fc.textContent = `Показано: ${filtered.length} из ${substitutionsJournal.length}`;
                }
            }

            renderSubstCalendar();
            renderSubstHistory();
        }

        function renderSubstCalendar() {
            const cal = document.getElementById("substCalendar");
            if (!cal) return;

            if (!substCalendarVisible) {
                cal.style.display = "none";
                return;
            }
            cal.style.display = "flex";

            if (!substCalMonth) {
                if (substViewMode === "day" && currentSubstDay) {
                    substCalMonth = currentSubstDay.slice(0, 7);
                } else if (currentSubstMonth) {
                    substCalMonth = currentSubstMonth;
                } else {
                    substCalMonth = new Date().toISOString().slice(0, 7);
                }
            }

            const [year, month] = substCalMonth.split("-").map(Number);
            const monthDate = new Date(year, month - 1, 1);

            const monthLabel = monthDate.toLocaleDateString("ru-RU", { month: "long", year: "numeric" });
            const labelEl = document.getElementById("substCalMonthLabel");
            if (labelEl) {
                labelEl.textContent = monthLabel.charAt(0).toUpperCase() + monthLabel.slice(1);
            }

            const countsByDate = {};
            substitutionsJournal.forEach(r => {
                if (r.date) {
                    countsByDate[r.date] = (countsByDate[r.date] || 0) + 1;
                }
            });

            const todayIso = new Date().toISOString().slice(0, 10);

            const grid = document.getElementById("substCalGrid");
            grid.innerHTML = "";

            const firstDayIdx = monthDate.getDay() === 0 ? 6 : monthDate.getDay() - 1;
            const daysInMonth = new Date(year, month, 0).getDate();
            const daysInPrevMonth = new Date(year, month - 1, 0).getDate();

            for (let i = firstDayIdx - 1; i >= 0; i--) {
                const dayNum = daysInPrevMonth - i;
                const cell = document.createElement("div");
                cell.className = "subst-cal-day other-month";
                cell.textContent = dayNum;
                grid.appendChild(cell);
            }

            for (let d = 1; d <= daysInMonth; d++) {
                const iso = `${year}-${String(month).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
                const cell = document.createElement("div");
                cell.className = "subst-cal-day";
                cell.dataset.date = iso;

                const dayOfWeek = new Date(year, month - 1, d).getDay();
                if (dayOfWeek === 0 || dayOfWeek === 6) cell.classList.add("weekend");

                const count = countsByDate[iso] || 0;
                if (count > 0) cell.classList.add("has-substitutions");
                if (iso === todayIso) cell.classList.add("today");
                if (iso === currentSubstDay && substViewMode === "day") cell.classList.add("selected");

                cell.innerHTML = `<span>${d}</span>${count > 0 ? `<span class="cal-count">${count}</span>` : ""}`;

                cell.addEventListener("click", () => {
                    currentSubstDay = iso;
                    const newMonth = iso.slice(0, 7);
                    if (newMonth !== currentSubstMonth) {
                        currentSubstMonth = newMonth;
                        const monthInp = document.getElementById("substMonthFilter");
                        if (monthInp) monthInp.value = newMonth;
                    }
                    if (substViewMode !== "day") {
                        substViewMode = "day";
                        const viewToggle = document.getElementById("substViewToggle");
                        if (viewToggle) {
                            viewToggle.querySelectorAll(".subst-view-btn").forEach(b => {
                                b.classList.toggle("active", b.dataset.view === "day");
                            });
                        }
                        const monthControls = document.getElementById("substMonthControls");
                        const dayControls = document.getElementById("substDayControls");
                        if (monthControls) monthControls.style.display = "none";
                        if (dayControls) dayControls.style.display = "inline-flex";
                    }
                    const dayInp = document.getElementById("substDayFilter");
                    if (dayInp) dayInp.value = iso;
                    renderSubstitutionsPanel();
                });

                grid.appendChild(cell);
            }

            const totalCells = firstDayIdx + daysInMonth;
            const remaining = (7 - (totalCells % 7)) % 7;
            for (let i = 1; i <= remaining; i++) {
                const cell = document.createElement("div");
                cell.className = "subst-cal-day other-month";
                cell.textContent = i;
                grid.appendChild(cell);
            }
        }

        function clearSubstForm() {
            ["substAbsentTeacher", "substAbsentSubject", "substSubstituteTeacher", "substSubstituteSubject"]
                .forEach(id => { const el = document.getElementById(id); if (el) el.value = ""; });
            const typeInp = document.getElementById("substLessonType");
            if (typeInp) typeInp.value = "Стандартный урок";
            const err = document.getElementById("substAddError");
            if (err) err.textContent = "";
            editingSubstId = null;
            const submitBtn = document.getElementById("substAddSubmitBtn");
            if (submitBtn) submitBtn.textContent = "➕ Добавить";
            const dateInp = document.getElementById("substDateInput");
            if (dateInp) {
                if (substViewMode === "day" && currentSubstDay) {
                    dateInp.value = currentSubstDay;
                } else {
                    dateInp.value = new Date().toISOString().slice(0, 10);
                }
            }
            updateSubstDateHint();
        }

        function submitSubstForm() {
            const dateInp = document.getElementById("substDateInput");
            const typeInp = document.getElementById("substLessonType");
            const absentTeacherInp = document.getElementById("substAbsentTeacher");
            const absentSubjectInp = document.getElementById("substAbsentSubject");
            const substituteTeacherInp = document.getElementById("substSubstituteTeacher");
            const substituteSubjectInp = document.getElementById("substSubstituteSubject");
            const err = document.getElementById("substAddError");

            const date = dateInp.value;
            const lessonType = typeInp.value || "Стандартный урок";
            const absentTeacher = absentTeacherInp.value.trim();
            const absentSubject = absentSubjectInp.value.trim();
            const substituteTeacher = substituteTeacherInp.value.trim();
            const substituteSubject = substituteSubjectInp.value.trim();

            if (err) err.textContent = "";
            if (!date) { if (err) err.textContent = "❌ Укажите дату"; return; }
            if (!lessonType) { if (err) err.textContent = "❌ Выберите тип урока"; return; }
            if (!absentTeacher) { if (err) err.textContent = "❌ Выберите отсутствующего учителя"; return; }
            if (!absentSubject) { if (err) err.textContent = "❌ Выберите предмет отсутствующего"; return; }
            if (!substituteTeacher) { if (err) err.textContent = "❌ Выберите заменяющего учителя"; return; }
            if (!substituteSubject) { if (err) err.textContent = "❌ Выберите предмет заменяющего"; return; }
            if (absentTeacher === substituteTeacher) { if (err) err.textContent = "❌ Заменяющий не может совпадать с отсутствующим"; return; }

            const day = getDayNameFromDate(new Date(date + "T00:00:00"));

            if (editingSubstId) {
                const r = substitutionsJournal.find(x => x.id === editingSubstId);
                if (!r) { editingSubstId = null; return; }
                const snapshotBefore = JSON.parse(JSON.stringify(r));
                r.date = date; r.day = day;
                r.lessonType = lessonType;
                r.absentTeacher = absentTeacher;
                r.absentSubject = absentSubject;
                r.substituteTeacher = substituteTeacher;
                r.substituteSubject = substituteSubject;
                addSubstHistoryEntry({
                    type: "edit",
                    recordId: r.id,
                    description: `${lessonType}: ${absentTeacher} (${absentSubject}) → ${substituteTeacher} (${substituteSubject}), ${date}`,
                    snapshotBefore
                });
                showToast("✓ Запись обновлена", "success");
            } else {
                const record = {
                    id: "subst_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8),
                    date, day,
                    lessonType,
                    absentTeacher, absentSubject,
                    substituteTeacher, substituteSubject,
                    note: "",
                    status: "confirmed",
                    createdAt: new Date().toISOString()
                };
                substitutionsJournal.unshift(record);
                addSubstHistoryEntry({
                    type: "add",
                    recordId: record.id,
                    description: `${lessonType}: ${absentTeacher} (${absentSubject}) → ${substituteTeacher} (${substituteSubject}), ${date}`
                });
                showToast("✓ Замена добавлена в журнал", "success");
            }

            persistSubstitutions();
            clearSubstForm();
            renderSubstitutionsPanel();
        }

        function startEditSubstRecord(id) {
            const r = substitutionsJournal.find(x => x.id === id);
            if (!r) return;
            editingSubstId = id;

            populateSubstFormSelects();

            document.getElementById("substDateInput").value = r.date || "";
            document.getElementById("substLessonType").value = r.lessonType || "Стандартный урок";
            document.getElementById("substAbsentTeacher").value = r.absentTeacher || "";
            document.getElementById("substAbsentSubject").value = r.absentSubject || "";
            document.getElementById("substSubstituteTeacher").value = r.substituteTeacher || "";
            document.getElementById("substSubstituteSubject").value = r.substituteSubject || "";

            const submitBtn = document.getElementById("substAddSubmitBtn");
            if (submitBtn) submitBtn.textContent = "💾 Сохранить изменения";
            const err = document.getElementById("substAddError");
            if (err) err.textContent = "";
            document.querySelector(".subst-add-panel").scrollIntoView({ behavior: "smooth", block: "start" });
        }

        function deleteSubstRecord(id) {
            const r = substitutionsJournal.find(x => x.id === id);
            if (!r) return;
            if (!confirm(`Удалить запись:\n${r.absentTeacher} (${r.absentSubject}) → ${r.substituteTeacher} (${r.substituteSubject}) на ${r.date}?`)) return;

            const idx = substitutionsJournal.findIndex(x => x.id === id);
            const snapshot = JSON.parse(JSON.stringify(r));
            substitutionsJournal.splice(idx, 1);
            addSubstHistoryEntry({
                type: "delete",
                recordId: id,
                index: idx,
                snapshot,
                description: `${r.absentTeacher} → ${r.substituteTeacher}, ${r.date}`
            });
            persistSubstitutions();
            renderSubstitutionsPanel();
            showToast("🗑 Запись удалена", "warning");
        }

        function initSubstitutionsTab() {
            const monthInp = document.getElementById("substMonthFilter");
            const dayInp = document.getElementById("substDayFilter");
            const searchInp = document.getElementById("substSearch");
            const addBtn = document.getElementById("substAddSubmitBtn");
            const clearBtn = document.getElementById("substClearFormBtn");
            const dateInp = document.getElementById("substDateInput");
            const exportBtn = document.getElementById("substExportExcelBtn");
            const printBtn = document.getElementById("substPrintBtn");
            const histBtn = document.getElementById("substHistoryToggleBtn");
            const closeHistBtn = document.getElementById("substCloseHistoryBtn");
            const undoLastBtn = document.getElementById("substUndoLastBtn");
            const calendarToggleBtn = document.getElementById("substCalendarToggleBtn");

            const viewToggle = document.getElementById("substViewToggle");
            const monthControls = document.getElementById("substMonthControls");
            const dayControls = document.getElementById("substDayControls");
            const prevDayBtn = document.getElementById("substPrevDayBtn");
            const nextDayBtn = document.getElementById("substNextDayBtn");
            const todayBtn = document.getElementById("substTodayBtn");

            const calPrevBtn = document.getElementById("substCalPrevBtn");
            const calNextBtn = document.getElementById("substCalNextBtn");

            const now = new Date();
            const ym = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
            if (!currentSubstMonth) currentSubstMonth = ym;
            if (monthInp) monthInp.value = currentSubstMonth;

            if (!currentSubstDay) currentSubstDay = new Date().toISOString().slice(0, 10);
            if (dayInp) dayInp.value = currentSubstDay;

            if (dateInp && !dateInp.value) dateInp.value = new Date().toISOString().slice(0, 10);

            if (viewToggle) {
                viewToggle.querySelectorAll(".subst-view-btn").forEach(btn => {
                    btn.addEventListener("click", () => {
                        const view = btn.dataset.view;
                        if (view === substViewMode) return;
                        substViewMode = view;

                        viewToggle.querySelectorAll(".subst-view-btn").forEach(b => b.classList.remove("active"));
                        btn.classList.add("active");

                        if (view === "day") {
                            monthControls.style.display = "none";
                            dayControls.style.display = "inline-flex";
                            if (!currentSubstDay) currentSubstDay = new Date().toISOString().slice(0, 10);
                            if (dayInp) dayInp.value = currentSubstDay;
                            substCalMonth = currentSubstDay.slice(0, 7);
                        } else {
                            monthControls.style.display = "inline-flex";
                            dayControls.style.display = "none";
                            if (currentSubstMonth) substCalMonth = currentSubstMonth;
                        }

                        renderSubstitutionsPanel();
                    });
                });
            }

            if (monthInp) monthInp.addEventListener("change", () => {
                currentSubstMonth = monthInp.value;
                if (substViewMode === "month") substCalMonth = currentSubstMonth;
                renderSubstitutionsPanel();
            });

            if (dayInp) dayInp.addEventListener("change", () => {
                currentSubstDay = dayInp.value;
                if (currentSubstDay) {
                    const newMonth = currentSubstDay.slice(0, 7);
                    if (newMonth !== currentSubstMonth) {
                        currentSubstMonth = newMonth;
                        if (monthInp) monthInp.value = newMonth;
                    }
                    substCalMonth = newMonth;
                }
                renderSubstitutionsPanel();
            });

            function shiftDay(deltaDays) {
                if (!currentSubstDay) currentSubstDay = new Date().toISOString().slice(0, 10);
                const d = new Date(currentSubstDay + "T00:00:00");
                d.setDate(d.getDate() + deltaDays);
                currentSubstDay = d.toISOString().slice(0, 10);
                if (dayInp) dayInp.value = currentSubstDay;
                const newMonth = currentSubstDay.slice(0, 7);
                if (newMonth !== currentSubstMonth) {
                    currentSubstMonth = newMonth;
                    if (monthInp) monthInp.value = newMonth;
                }
                substCalMonth = newMonth;
                renderSubstitutionsPanel();
            }

            if (prevDayBtn) prevDayBtn.addEventListener("click", () => shiftDay(-1));
            if (nextDayBtn) nextDayBtn.addEventListener("click", () => shiftDay(1));
            if (todayBtn) todayBtn.addEventListener("click", () => {
                currentSubstDay = new Date().toISOString().slice(0, 10);
                if (dayInp) dayInp.value = currentSubstDay;
                const newMonth = currentSubstDay.slice(0, 7);
                if (newMonth !== currentSubstMonth) {
                    currentSubstMonth = newMonth;
                    if (monthInp) monthInp.value = newMonth;
                }
                substCalMonth = newMonth;
                renderSubstitutionsPanel();
            });

            if (searchInp) searchInp.addEventListener("input", () => {
                currentSubstSearch = searchInp.value.trim();
                renderSubstitutionsPanel();
            });

            if (addBtn) addBtn.addEventListener("click", submitSubstForm);
            if (clearBtn) clearBtn.addEventListener("click", clearSubstForm);
            if (exportBtn) exportBtn.addEventListener("click", exportSubstitutionsToExcel);
            if (printBtn) printBtn.addEventListener("click", printSubstitutions);

            if (calendarToggleBtn) {
                calendarToggleBtn.addEventListener("click", () => {
                    substCalendarVisible = !substCalendarVisible;
                    calendarToggleBtn.classList.toggle("btn-primary", substCalendarVisible);
                    calendarToggleBtn.classList.toggle("btn-outline", !substCalendarVisible);
                    if (substCalendarVisible) {
                        if (substViewMode === "day" && currentSubstDay) {
                            substCalMonth = currentSubstDay.slice(0, 7);
                        } else if (currentSubstMonth) {
                            substCalMonth = currentSubstMonth;
                        }
                    }
                    renderSubstCalendar();
                });
            }

            if (calPrevBtn) calPrevBtn.addEventListener("click", () => {
                if (!substCalMonth) return;
                const [y, m] = substCalMonth.split("-").map(Number);
                const d = new Date(y, m - 2, 1);
                substCalMonth = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
                renderSubstCalendar();
            });
            if (calNextBtn) calNextBtn.addEventListener("click", () => {
                if (!substCalMonth) return;
                const [y, m] = substCalMonth.split("-").map(Number);
                const d = new Date(y, m, 1);
                substCalMonth = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
                renderSubstCalendar();
            });

            if (histBtn) histBtn.addEventListener("click", () => {
                substHistoryExpanded = !substHistoryExpanded;
                document.getElementById("substHistoryPanel").classList.toggle("visible", substHistoryExpanded);
                renderSubstHistory();
            });
            if (closeHistBtn) closeHistBtn.addEventListener("click", () => {
                substHistoryExpanded = false;
                document.getElementById("substHistoryPanel").classList.remove("visible");
            });
            if (undoLastBtn) undoLastBtn.addEventListener("click", undoLastSubstChange);

            ["substLessonType", "substAbsentTeacher", "substAbsentSubject", "substSubstituteTeacher", "substSubstituteSubject"]
                .forEach(id => {
                    const el = document.getElementById(id);
                    if (el) el.addEventListener("keydown", e => { if (e.key === "Enter") submitSubstForm(); });
                });

            if (substViewMode === "day") {
                if (monthControls) monthControls.style.display = "none";
                if (dayControls) dayControls.style.display = "inline-flex";
                if (viewToggle) {
                    viewToggle.querySelectorAll(".subst-view-btn").forEach(b => {
                        b.classList.toggle("active", b.dataset.view === "day");
                    });
                }
                if (dayInp) dayInp.value = currentSubstDay;
            } else {
                if (monthControls) monthControls.style.display = "inline-flex";
                if (dayControls) dayControls.style.display = "none";
            }

            initDictManage();
        }

        function initDictManage() {
            document.getElementById("substManageDictionariesBtn").addEventListener("click", openDictManage);
            document.getElementById("closeDictManageBtn").addEventListener("click", closeDictManage);
            document.getElementById("dictManageModal").addEventListener("click", e => {
                if (e.target.id === "dictManageModal") closeDictManage();
            });

            document.querySelectorAll(".console-tab[data-dict-tab]").forEach(tab => {
                tab.addEventListener("click", () => {
                    document.querySelectorAll(".console-tab[data-dict-tab]").forEach(t => t.classList.remove("active"));
                    document.querySelectorAll(".dict-content").forEach(c => c.classList.remove("active"));
                    tab.classList.add("active");
                    document.querySelector(`.dict-content[data-dict-tab="${tab.dataset.dictTab}"]`).classList.add("active");
                });
            });

            const addTeacherBtn = document.getElementById("dictAddTeacherBtn");
            const addSubjectBtn = document.getElementById("dictAddSubjectBtn");
            const teacherInput = document.getElementById("dictTeacherInput");
            const subjectInput = document.getElementById("dictSubjectInput");

            addTeacherBtn.addEventListener("click", () => {
                const name = teacherInput.value.trim();
                if (!name) return;
                const all = getAllTeachersForJournal();
                if (all.includes(name)) { showToast("Такой учитель уже есть", "warning"); return; }
                dictTeachersManual.push(name);
                persistDictTeachers();
                teacherInput.value = "";
                renderDictLists();
                populateSubstFormSelects();
                showToast(`✓ Добавлен: ${name}`, "success");
            });
            teacherInput.addEventListener("keydown", e => { if (e.key === "Enter") addTeacherBtn.click(); });

            addSubjectBtn.addEventListener("click", () => {
                const name = subjectInput.value.trim();
                if (!name) return;
                const all = getAllSubjectsForJournal();
                if (all.includes(name)) { showToast("Такой предмет уже есть", "warning"); return; }
                dictSubjectsManual.push(name);
                persistDictSubjects();
                subjectInput.value = "";
                renderDictLists();
                populateSubstFormSelects();
                showToast(`✓ Добавлен: ${name}`, "success");
            });
            subjectInput.addEventListener("keydown", e => { if (e.key === "Enter") addSubjectBtn.click(); });
        }

        function openDictManage() {
            renderDictLists();
            document.getElementById("dictManageModal").classList.add("visible");
        }

        function closeDictManage() {
            document.getElementById("dictManageModal").classList.remove("visible");
        }

        function renderDictLists() {
            const tList = document.getElementById("dictTeacherList");
            const autoTeachers = new Set(scheduleData.map(t => t.name).filter(Boolean));
            const allTeachers = getAllTeachersForJournal();

            if (allTeachers.length === 0) {
                tList.innerHTML = '<div class="dict-empty">Пока никого нет. Добавьте первого учителя.</div>';
            } else {
                tList.innerHTML = "";
                allTeachers.forEach(name => {
                    const isAuto = autoTeachers.has(name);
                    const item = document.createElement("div");
                    item.className = "dict-item " + (isAuto ? "auto" : "manual");
                    item.innerHTML = `
                        <span class="dict-name">${escapeHtml(name)}</span>
                        <span class="dict-badge">${isAuto ? "из расписания" : "добавлен"}</span>
                    `;
                    if (!isAuto) {
                        const delBtn = document.createElement("button");
                        delBtn.className = "icon-btn danger";
                        delBtn.title = "Удалить";
                        delBtn.textContent = "🗑";
                        delBtn.addEventListener("click", () => {
                            if (!confirm(`Удалить «${name}» из справочника?`)) return;
                            dictTeachersManual = dictTeachersManual.filter(n => n !== name);
                            persistDictTeachers();
                            renderDictLists();
                            populateSubstFormSelects();
                            showToast("🗑 Удалено", "warning");
                        });
                        item.appendChild(delBtn);
                    } else {
                        item.innerHTML += '<span style="font-size:0.65rem; color:var(--text-muted);">управляется в «Учительском»</span>';
                    }
                    tList.appendChild(item);
                });
            }

            const sList = document.getElementById("dictSubjectList");
            const autoSubjects = new Set();
            classesData.forEach(c => {
                DAYS.forEach(day => {
                    (c.lessons[day] || []).forEach(l => {
                        const txt = (normalizeLesson(l).text || "").trim();
                        if (txt) autoSubjects.add(txt);
                    });
                });
            });
            const allSubjects = getAllSubjectsForJournal();

            if (allSubjects.length === 0) {
                sList.innerHTML = '<div class="dict-empty">Пока ничего нет.</div>';
            } else {
                sList.innerHTML = "";
                allSubjects.forEach(name => {
                    const isAuto = autoSubjects.has(name);
                    const item = document.createElement("div");
                    item.className = "dict-item " + (isAuto ? "auto" : "manual");
                    item.innerHTML = `
                        <span class="dict-name">${escapeHtml(name)}</span>
                        <span class="dict-badge">${isAuto ? "из «Детского»" : "добавлен"}</span>
                    `;
                    if (!isAuto) {
                        const delBtn = document.createElement("button");
                        delBtn.className = "icon-btn danger";
                        delBtn.title = "Удалить";
                        delBtn.textContent = "🗑";
                        delBtn.addEventListener("click", () => {
                            if (!confirm(`Удалить «${name}» из справочника?`)) return;
                            dictSubjectsManual = dictSubjectsManual.filter(n => n !== name);
                            persistDictSubjects();
                            renderDictLists();
                            populateSubstFormSelects();
                            showToast("🗑 Удалено", "warning");
                        });
                        item.appendChild(delBtn);
                    } else {
                        item.innerHTML += '<span style="font-size:0.65rem; color:var(--text-muted);">из «Детского»</span>';
                    }
                    sList.appendChild(item);
                });
            }
        }

        function exportSubstitutionsToExcel() {
            try {
                const wb = XLSX.utils.book_new();
                const filtered = getSubstFiltered();

                const rows = [["Дата", "День", "Тип урока", "Отсутствующий учитель", "Его предмет", "Заменяющий учитель", "Его предмет"]];
                filtered.slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).forEach(r => {
                    rows.push([
                        r.date || "",
                        DAY_SHORT[r.day] || r.day || "",
                        r.lessonType || "Стандартный урок",
                        r.absentTeacher || "",
                        r.absentSubject || "",
                        r.substituteTeacher || "",
                        r.substituteSubject || ""
                    ]);
                });
                const ws = XLSX.utils.aoa_to_sheet(rows);
                ws["!cols"] = [{ wch: 12 }, { wch: 6 }, { wch: 22 }, { wch: 24 }, { wch: 22 }, { wch: 24 }, { wch: 22 }];
                XLSX.utils.book_append_sheet(wb, ws, "Журнал замен");
                let fileSuffix = formatDateForFile();
                if (substViewMode === "day" && currentSubstDay) {
                    fileSuffix = currentSubstDay;
                } else if (substViewMode === "month" && currentSubstMonth) {
                    fileSuffix = currentSubstMonth;
                }
                XLSX.writeFile(wb, `zhurnal_zamen_${fileSuffix}.xlsx`);
                showToast("📤 Excel-журнал скачан", "success");
            } catch (err) {
                console.error(err);
                showToast(`❌ Ошибка экспорта: ${err.message}`, "error");
            }
        }

        function printSubstitutions() {
            const filtered = getSubstFiltered();

            let title;
            if (substViewMode === "day" && currentSubstDay) {
                const d = new Date(currentSubstDay + "T00:00:00");
                const dayName = getDayNameFromDate(d);
                const dayShort = DAY_SHORT[dayName] || dayName;
                const dateStr = d.toLocaleDateString("ru-RU", { day: "numeric", month: "long", year: "numeric" });
                title = `Журнал замен учителей на ${dayShort}, ${dateStr}`;
            } else if (substViewMode === "month" && currentSubstMonth) {
                const [y, m] = currentSubstMonth.split("-");
                const monthName = new Date(parseInt(y), parseInt(m) - 1, 1)
                    .toLocaleDateString("ru-RU", { month: "long", year: "numeric" });
                title = `Журнал замен учителей за ${monthName}`;
            } else {
                title = `Журнал замен учителей за всё время`;
            }

            const win = window.open("", "_blank");
            if (!win) { showToast("❌ Разрешите всплывающие окна", "error"); return; }

            let html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Журнал замен</title>
            <style>
                body { font-family: 'Times New Roman', serif; padding: 20px; color: #000; }
                h1 { font-size: 16pt; text-align: center; margin-bottom: 6px; }
                h2 { font-size: 12pt; text-align: center; font-weight: normal; margin-bottom: 20px; color: #444; }
                table { width: 100%; border-collapse: collapse; font-size: 10pt; }
                th { background: #1a3a6b; color: white; padding: 6px; border: 1px solid #333; text-align: left; vertical-align: middle; }
                td { padding: 5px 6px; border: 1px solid #999; vertical-align: top; }
                tr:nth-child(even) td { background: #f5f8fc; }
                .sign-col { width: 120px; min-width: 100px; }
                .sign-cell { height: 24px; }
                @media print { body { padding: 0; } }
            </style></head><body>`;
            html += `<h1>МОУ «Северная СОШ №2»</h1>`;
            html += `<h2>${escapeHtml(title)}</h2>`;

            if (filtered.length === 0) {
                html += `<div style="text-align:center; padding:30px; font-style:italic; color:#888;">Записей за указанный период нет</div>`;
            } else {
                html += `<table><thead><tr>
                    <th>Дата</th>
                    <th>День</th>
                    <th>Тип урока</th>
                    <th>Отсутствующий</th>
                    <th>Его предмет</th>
                    <th>Заменяющий</th>
                    <th>Его предмет</th>
                    <th class="sign-col">Подпись</th>
                </tr></thead><tbody>`;

                filtered.slice().sort((a, b) => {
                    if ((a.date || "") !== (b.date || "")) return (a.date || "").localeCompare(b.date || "");
                    return (a.createdAt || "").localeCompare(b.createdAt || "");
                }).forEach(r => {
                    html += `<tr>
                        <td>${r.date ? new Date(r.date).toLocaleDateString("ru-RU") : ""}</td>
                        <td>${escapeHtml(DAY_SHORT[r.day] || r.day || "")}</td>
                        <td>${escapeHtml(r.lessonType || "Стандартный урок")}</td>
                        <td>${escapeHtml(r.absentTeacher || "")}</td>
                        <td>${escapeHtml(r.absentSubject || "")}</td>
                        <td>${escapeHtml(r.substituteTeacher || "")}</td>
                        <td>${escapeHtml(r.substituteSubject || "")}</td>
                        <td class="sign-cell"></td>
                    </tr>`;
                });
                html += `</tbody></table>`;
            }

            html += `<div style="margin-top:20px; font-size:9pt; color:#666;">Сформировано: ${new Date().toLocaleString("ru-RU")}</div>`;
            html += `</body></html>`;

            win.document.write(html);
            win.document.close();
            setTimeout(() => win.print(), 300);
        }

        // ==================== КОНСТРУКТОР: ЛОГИКА ====================
        let constructorTabUnlockedUntil = 0;

        function loadConstructorConfig() {
            try {
                const saved = localStorage.getItem(STORAGE_KEYS.constructor);
                if (saved) {
                    const parsed = JSON.parse(saved);
                    if (parsed && typeof parsed === "object") {
                        return deepMerge(JSON.parse(JSON.stringify(FACTORY_CONFIG)), parsed);
                    }
                }
            } catch (e) { console.error(e); }
            return JSON.parse(JSON.stringify(FACTORY_CONFIG));
        }

        function persistConstructorConfig() {
            try {
                localStorage.setItem(STORAGE_KEYS.constructor, JSON.stringify(constructorConfig));
            } catch (e) { console.error(e); }
        }

        function deepMerge(target, source) {
            Object.keys(source).forEach(key => {
                const sv = source[key];
                if (sv && typeof sv === "object" && !Array.isArray(sv)) {
                    if (!target[key] || typeof target[key] !== "object") target[key] = {};
                    deepMerge(target[key], sv);
                } else if (Array.isArray(sv)) {
                    target[key] = sv;
                } else if (sv !== undefined && sv !== null && sv !== "") {
                    target[key] = sv;
                }
            });
            return target;
        }

        function applyConstructorConfig() {
            if (!constructorConfig) return;
            const c = constructorConfig;

            const h1 = document.querySelector(".school-title h1");
            const subP = document.querySelector(".school-title p");
            const noteDiv = document.querySelector(".school-title .subtitle");
            if (h1) {
                h1.textContent = c.header.title || "";
                h1.style.display = c.header.showTitle ? "" : "none";
                h1.style.color = c.header.titleColor || "";
            }
            if (subP) {
                subP.textContent = c.header.subtitle || "";
                subP.style.display = c.header.showSubtitle ? "" : "none";
            }
            if (noteDiv) {
                noteDiv.textContent = c.header.note || "";
                noteDiv.style.display = c.header.showNote ? "" : "none";
            }

            renderScheduleTabsFromConfig();

            const root = document.documentElement;
            if (c.theme.accent) root.style.setProperty("--accent", c.theme.accent);
            if (c.theme.bgPage) root.style.setProperty("--bg-page", c.theme.bgPage);
            if (c.theme.bgContainer) root.style.setProperty("--bg-container", c.theme.bgContainer);
            if (c.theme.textPrimary) root.style.setProperty("--text-primary", c.theme.textPrimary);
            if (c.theme.textSecondary) root.style.setProperty("--text-secondary", c.theme.textSecondary);
            if (c.theme.border) root.style.setProperty("--border", c.theme.border);
            if (c.theme.substBg) root.style.setProperty("--substitute-bg", c.theme.substBg);
            if (c.theme.substText) root.style.setProperty("--substitute-text", c.theme.substText);
            if (c.theme.success) root.style.setProperty("--success", c.theme.success);
            if (c.theme.danger) root.style.setProperty("--danger", c.theme.danger);

            const appContainer = document.querySelector(".app-container");
            if (appContainer) {
                appContainer.style.maxWidth = c.sizes.containerWidth + "px";
                appContainer.style.padding = "16px " + c.sizes.containerPadding + "px 14px";
            }
            if (h1) h1.style.fontSize = c.sizes.fontSizeTitle + "rem";

            root.style.setProperty("--constr-font-base", c.sizes.fontSizeBase + "rem");
            root.style.setProperty("--constr-radius", c.sizes.radius + "px");
            root.style.setProperty("--constr-row-padding", c.sizes.rowPadding + "px");

            let styleEl = document.getElementById("constructorDynamicStyle");
            if (!styleEl) {
                styleEl = document.createElement("style");
                styleEl.id = "constructorDynamicStyle";
                document.head.appendChild(styleEl);
            }
            styleEl.textContent = `
                .lesson-cell, .subst-table td, .admin-table td {
                    font-size: var(--constr-font-base, 0.78rem) !important;
                }
                .subst-section, .app-container, .modal, .schedule-tabs,
                .filter-bar, .table-wrapper, .subst-add-panel, .subst-dashboard > * {
                    border-radius: var(--constr-radius, 12px) !important;
                }
                .lesson-cell, .subst-table td, td {
                    padding-top: var(--constr-row-padding, 6px) !important;
                    padding-bottom: var(--constr-row-padding, 6px) !important;
                }
                .app-container {
                    backdrop-filter: ${c.effects.blur ? "blur(8px)" : "none"} !important;
                    box-shadow: ${c.effects.shadow ? "var(--shadow-container)" : "none"} !important;
                    transition: ${c.effects.transitions ? "background 0.3s ease, box-shadow 0.3s ease" : "none"} !important;
                }
            `;

            const animatedBg = document.getElementById("animatedBg");
            if (animatedBg) {
                animatedBg.style.display = c.effects.animatedBg ? "" : "none";
                animatedBg.style.opacity = c.effects.bgIntensity;
            }

            const btnEdit = document.getElementById("editToggleBtn");
            const btnMySchedule = document.getElementById("myScheduleBtn");
            const btnConsole = document.querySelector(".console-btn-wrapper");
            const btnTheme = document.querySelector(".theme-toggle");
            const btnMobile = document.getElementById("mobileToggle");
            const indicators = document.querySelectorAll(".edit-indicator, .substitute-indicator");
            const adminIndicator = document.getElementById("adminIndicator");
            const logoutBtn = document.getElementById("logoutAdminBtn");

            if (btnEdit) btnEdit.style.display = c.buttons.edit ? "" : "none";
            if (btnMySchedule) btnMySchedule.style.display = c.buttons.mySchedule ? "" : "none";
            if (btnConsole) btnConsole.style.display = c.buttons.console ? "" : "none";
            if (btnTheme) btnTheme.style.display = c.buttons.theme ? "" : "none";
            if (btnMobile) btnMobile.style.display = c.buttons.mobile !== false ? "" : "none";
            indicators.forEach(el => { el.style.visibility = c.buttons.indicators ? "" : "hidden"; });
            if (adminIndicator) adminIndicator.style.visibility = c.buttons.adminIndicator ? "" : "hidden";
            if (logoutBtn && !c.buttons.adminIndicator) logoutBtn.style.display = "none";

            const footerHint = document.getElementById("footerHint");
            const statusBadge = document.getElementById("statusBadge");
            const historyBtn = document.getElementById("showHistoryBtn");
            if (footerHint) footerHint.textContent = c.footer.hint || "";
            if (statusBadge) statusBadge.style.display = c.footer.showStatus ? "" : "none";
            if (historyBtn) historyBtn.style.display = c.footer.showHistory ? "" : "none";

            renderCustomTabs();
            bindCustomTabClicks();
        }

        function renderScheduleTabsFromConfig() {
            const tabsContainer = document.getElementById("scheduleTabs");
            if (!tabsContainer) return;

            tabsContainer.querySelectorAll(".schedule-tab").forEach(t => {
                const key = t.dataset.schedule;
                const cfgTab = constructorConfig.tabs.find(x => x.key === key);
                if (!cfgTab || !cfgTab.visible) t.remove();
            });

            constructorConfig.tabs.forEach((tab) => {
                if (!tab.visible) return;
                let btn = tabsContainer.querySelector(`.schedule-tab[data-schedule="${tab.key}"]`);
                if (!btn) {
                    btn = document.createElement("button");
                    btn.className = "schedule-tab";
                    btn.dataset.schedule = tab.key;
                    btn.innerHTML = `<span class="tab-icon">${escapeHtml(tab.icon || "📄")}</span><span class="tab-label">${escapeHtml(tab.label)}</span>`;
                    if (tab.key.startsWith("custom:")) {
                        btn.addEventListener("click", () => {
                            const tabId = tab.key.split(":")[1];
                            switchToCustomTab(tabId);
                        });
                    } else {
                        btn.addEventListener("click", () => switchSchedule(tab.key));
                    }
                    tabsContainer.appendChild(btn);
                } else {
                    const iconEl = btn.querySelector(".tab-icon");
                    const labelEl = btn.querySelector(".tab-label");
                    if (iconEl) iconEl.textContent = tab.icon || "📄";
                    if (labelEl) labelEl.textContent = tab.label;
                }
            });

            tabsContainer.querySelectorAll(".schedule-tab").forEach(t => {
                t.classList.toggle("active", t.dataset.schedule === currentScheduleType);
            });
        }

        function renderCustomTabs() {
            document.querySelectorAll(".custom-tab-panel").forEach(p => p.remove());

            const appContainer = document.querySelector(".app-container");
            const footer = appContainer.querySelector(".footer-note");
            if (!appContainer || !footer) return;

            constructorConfig.customTabs.forEach(tab => {
                const panel = document.createElement("div");
                panel.className = "custom-tab-panel";
                panel.dataset.customTab = tab.id;
                panel.innerHTML = `<div class="custom-content-box">${tab.content || "<p>Пустая вкладка</p>"}</div>`;
                appContainer.insertBefore(panel, footer);
            });
        }

        function switchToCustomTab(tabId) {
            currentScheduleType = "custom:" + tabId;
            localStorage.setItem(STORAGE_KEYS.scheduleType, "custom:" + tabId);

            document.body.classList.remove("mode-teachers", "mode-classes", "mode-iup", "mode-substitutions");
            document.body.classList.add("mode-custom");

            document.querySelectorAll(".schedule-tab").forEach(t => {
                t.classList.toggle("active", t.dataset.schedule === "custom:" + tabId);
            });

            document.querySelectorAll(".custom-tab-panel").forEach(p => {
                p.classList.toggle("active", p.dataset.customTab === tabId);
            });

            const hintText = document.getElementById("scheduleHintText");
            if (hintText) hintText.innerHTML = "<strong>Пользовательская вкладка</strong>";
        }

        function bindCustomTabClicks() {
            document.querySelectorAll(".schedule-tab").forEach(btn => {
                const key = btn.dataset.schedule;
                if (key && key.startsWith("custom:") && !btn._boundCustom) {
                    btn._boundCustom = true;
                    btn.addEventListener("click", () => {
                        const tabId = key.split(":")[1];
                        switchToCustomTab(tabId);
                    });
                }
            });
        }

        function setConstrStatus(msg, type) {
            const el = document.getElementById("constrStatus");
            if (!el) return;
            el.textContent = msg;
            el.className = "constr-status" + (type ? " " + type : "");
            clearTimeout(el._timer);
            el._timer = setTimeout(() => {
                el.textContent = "";
                el.className = "constr-status";
            }, 4000);
        }

        function fillConstructorFields() {
            const c = constructorConfig;

            setVal("constrHeaderTitle", c.header.title);
            setVal("constrHeaderSubtitle", c.header.subtitle);
            setVal("constrHeaderNote", c.header.note);
            setChk("constrShowHeaderTitle", c.header.showTitle);
            setChk("constrShowHeaderSubtitle", c.header.showSubtitle);
            setChk("constrShowHeaderNote", c.header.showNote);
            setColorFields("constrHeaderTitleColor", "constrHeaderTitleColorHex", c.header.titleColor);

            setColorFields("constrColorAccent", "constrColorAccentHex", c.theme.accent);
            setColorFields("constrColorBgPage", "constrColorBgPageHex", c.theme.bgPage);
            setColorFields("constrColorBgContainer", "constrColorBgContainerHex", c.theme.bgContainer);
            setColorFields("constrColorTextPrimary", "constrColorTextPrimaryHex", c.theme.textPrimary);
            setColorFields("constrColorTextSecondary", "constrColorTextSecondaryHex", c.theme.textSecondary);
            setColorFields("constrColorBorder", "constrColorBorderHex", c.theme.border);
            setColorFields("constrColorSubstBg", "constrColorSubstBgHex", c.theme.substBg);
            setColorFields("constrColorSubstText", "constrColorSubstTextHex", c.theme.substText);
            setColorFields("constrColorSuccess", "constrColorSuccessHex", c.theme.success);
            setColorFields("constrColorDanger", "constrColorDangerHex", c.theme.danger);

            const sz = c.sizes;
            setRange("constrSizeBase", "constrSizeBaseVal", sz.fontSizeBase, 2);
            setRange("constrSizeTitle", "constrSizeTitleVal", sz.fontSizeTitle, 2);
            setRange("constrRadius", "constrRadiusVal", sz.radius, 0);
            setRange("constrContainerWidth", "constrContainerWidthVal", sz.containerWidth, 0);
            setRange("constrContainerPadding", "constrContainerPaddingVal", sz.containerPadding, 0);
            setRange("constrRowPadding", "constrRowPaddingVal", sz.rowPadding, 0);

            setChk("constrEffectBg", c.effects.animatedBg);
            setChk("constrEffectShadow", c.effects.shadow);
            setChk("constrEffectBlur", c.effects.blur);
            setChk("constrEffectTransitions", c.effects.transitions);
            setRange("constrBgIntensity", "constrBgIntensityVal", c.effects.bgIntensity, 1);

            setChk("constrBtnEdit", c.buttons.edit);
            setChk("constrBtnMySchedule", c.buttons.mySchedule);
            setChk("constrBtnConsole", c.buttons.console);
            setChk("constrBtnTheme", c.buttons.theme);
            setChk("constrBtnMobile", c.buttons.mobile !== false);
            setChk("constrBtnIndicators", c.buttons.indicators);
            setChk("constrBtnAdminIndicator", c.buttons.adminIndicator);

            setVal("constrFooterHint", c.footer.hint);
            setChk("constrShowFooterStatus", c.footer.showStatus);
            setChk("constrShowFooterHistory", c.footer.showHistory);
        }

        function setVal(id, value) {
            const el = document.getElementById(id);
            if (el) el.value = value != null ? value : "";
        }
        function setChk(id, value) {
            const el = document.getElementById(id);
            if (el) el.checked = !!value;
        }
        function setRange(id, valId, value, decimals) {
            const el = document.getElementById(id);
            const valEl = document.getElementById(valId);
            if (el) el.value = value;
            if (valEl) valEl.textContent = Number(value).toFixed(decimals);
        }
        function setColorFields(colorId, hexId, value) {
            const colorEl = document.getElementById(colorId);
            const hexEl = document.getElementById(hexId);
            if (colorEl && isHex(value)) colorEl.value = value;
            if (hexEl) hexEl.value = value || "";
        }
        function isHex(v) { return typeof v === "string" && /^#[0-9a-fA-F]{6}$/.test(v); }

        function bindConstructorInputs() {
            const colorPairs = [
                ["constrHeaderTitleColor", "constrHeaderTitleColorHex"],
                ["constrColorAccent", "constrColorAccentHex"],
                ["constrColorBgPage", "constrColorBgPageHex"],
                ["constrColorBgContainer", "constrColorBgContainerHex"],
                ["constrColorTextPrimary", "constrColorTextPrimaryHex"],
                ["constrColorTextSecondary", "constrColorTextSecondaryHex"],
                ["constrColorBorder", "constrColorBorderHex"],
                ["constrColorSubstBg", "constrColorSubstBgHex"],
                ["constrColorSubstText", "constrColorSubstTextHex"],
                ["constrColorSuccess", "constrColorSuccessHex"],
                ["constrColorDanger", "constrColorDangerHex"]
            ];
            colorPairs.forEach(([cid, hid]) => {
                const c = document.getElementById(cid);
                const h = document.getElementById(hid);
                if (!c || !h) return;
                c.addEventListener("input", () => { h.value = c.value; });
                h.addEventListener("input", () => { if (isHex(h.value)) c.value = h.value; });
            });

            const rangePairs = [
                ["constrSizeBase", "constrSizeBaseVal", 2],
                ["constrSizeTitle", "constrSizeTitleVal", 2],
                ["constrRadius", "constrRadiusVal", 0],
                ["constrContainerWidth", "constrContainerWidthVal", 0],
                ["constrContainerPadding", "constrContainerPaddingVal", 0],
                ["constrRowPadding", "constrRowPaddingVal", 0],
                ["constrBgIntensity", "constrBgIntensityVal", 1]
            ];
            rangePairs.forEach(([id, valId, dec]) => {
                const el = document.getElementById(id);
                const valEl = document.getElementById(valId);
                if (!el || !valEl) return;
                el.addEventListener("input", () => {
                    valEl.textContent = Number(el.value).toFixed(dec);
                });
            });
        }

        function collectConstructorFields() {
            const c = constructorConfig;

            c.header.title = document.getElementById("constrHeaderTitle").value;
            c.header.subtitle = document.getElementById("constrHeaderSubtitle").value;
            c.header.note = document.getElementById("constrHeaderNote").value;
            c.header.showTitle = document.getElementById("constrShowHeaderTitle").checked;
            c.header.showSubtitle = document.getElementById("constrShowHeaderSubtitle").checked;
            c.header.showNote = document.getElementById("constrShowHeaderNote").checked;
            c.header.titleColor = document.getElementById("constrHeaderTitleColorHex").value || c.header.titleColor;

            c.theme.accent = document.getElementById("constrColorAccentHex").value || c.theme.accent;
            c.theme.bgPage = document.getElementById("constrColorBgPageHex").value || c.theme.bgPage;
            c.theme.bgContainer = document.getElementById("constrColorBgContainerHex").value || c.theme.bgContainer;
            c.theme.textPrimary = document.getElementById("constrColorTextPrimaryHex").value || c.theme.textPrimary;
            c.theme.textSecondary = document.getElementById("constrColorTextSecondaryHex").value || c.theme.textSecondary;
            c.theme.border = document.getElementById("constrColorBorderHex").value || c.theme.border;
            c.theme.substBg = document.getElementById("constrColorSubstBgHex").value || c.theme.substBg;
            c.theme.substText = document.getElementById("constrColorSubstTextHex").value || c.theme.substText;
            c.theme.success = document.getElementById("constrColorSuccessHex").value || c.theme.success;
            c.theme.danger = document.getElementById("constrColorDangerHex").value || c.theme.danger;

            c.sizes.fontSizeBase = parseFloat(document.getElementById("constrSizeBase").value);
            c.sizes.fontSizeTitle = parseFloat(document.getElementById("constrSizeTitle").value);
            c.sizes.radius = parseInt(document.getElementById("constrRadius").value);
            c.sizes.containerWidth = parseInt(document.getElementById("constrContainerWidth").value);
            c.sizes.containerPadding = parseInt(document.getElementById("constrContainerPadding").value);
            c.sizes.rowPadding = parseInt(document.getElementById("constrRowPadding").value);

            c.effects.animatedBg = document.getElementById("constrEffectBg").checked;
            c.effects.shadow = document.getElementById("constrEffectShadow").checked;
            c.effects.blur = document.getElementById("constrEffectBlur").checked;
            c.effects.transitions = document.getElementById("constrEffectTransitions").checked;
            c.effects.bgIntensity = parseFloat(document.getElementById("constrBgIntensity").value);

            c.buttons.edit = document.getElementById("constrBtnEdit").checked;
            c.buttons.mySchedule = document.getElementById("constrBtnMySchedule").checked;
            c.buttons.console = document.getElementById("constrBtnConsole").checked;
            c.buttons.theme = document.getElementById("constrBtnTheme").checked;
            c.buttons.mobile = document.getElementById("constrBtnMobile").checked;
            c.buttons.indicators = document.getElementById("constrBtnIndicators").checked;
            c.buttons.adminIndicator = document.getElementById("constrBtnAdminIndicator").checked;

            c.footer.hint = document.getElementById("constrFooterHint").value;
            c.footer.showStatus = document.getElementById("constrShowFooterStatus").checked;
            c.footer.showHistory = document.getElementById("constrShowFooterHistory").checked;
        }

        function renderConstructorTabsList() {
            const container = document.getElementById("constrTabsList");
            if (!container) return;
            container.innerHTML = "";

            constructorConfig.tabs.forEach((tab, idx) => {
                const row = document.createElement("div");
                row.className = "constr-tab-row";
                row.innerHTML = `
                    <div class="constr-tab-idx">${idx + 1}</div>
                    <input type="text" value="${escapeHtml(tab.label)}" data-field="label" data-idx="${idx}" placeholder="Название">
                    <input type="text" value="${escapeHtml(tab.icon)}" data-field="icon" data-idx="${idx}" maxlength="4" placeholder="🎯">
                    <div class="constr-tab-actions">
                        <label class="constr-check" title="Показать/скрыть" style="margin-right:4px;">
                            <input type="checkbox" data-field="visible" data-idx="${idx}" ${tab.visible ? "checked" : ""}>
                        </label>
                        <button class="icon-btn" data-act="up" data-idx="${idx}" title="Вверх">▲</button>
                        <button class="icon-btn" data-act="down" data-idx="${idx}" title="Вниз">▼</button>
                        ${tab.key === "teachers" ? "" : `<button class="icon-btn danger" data-act="del" data-idx="${idx}" title="Удалить">🗑</button>`}
                    </div>
                `;
                container.appendChild(row);
            });

            container.querySelectorAll("input[data-field]").forEach(inp => {
                inp.addEventListener("change", () => {
                    const idx = parseInt(inp.dataset.idx);
                    const field = inp.dataset.field;
                    if (field === "visible") {
                        constructorConfig.tabs[idx].visible = inp.checked;
                    } else {
                        constructorConfig.tabs[idx][field] = inp.value;
                    }
                    persistConstructorConfig();
                    applyConstructorConfig();
                });
            });

            container.querySelectorAll("[data-act]").forEach(btn => {
                btn.addEventListener("click", () => {
                    const idx = parseInt(btn.dataset.idx);
                    const act = btn.dataset.act;
                    if (act === "up" && idx > 0) {
                        [constructorConfig.tabs[idx - 1], constructorConfig.tabs[idx]] = [constructorConfig.tabs[idx], constructorConfig.tabs[idx - 1]];
                    } else if (act === "down" && idx < constructorConfig.tabs.length - 1) {
                        [constructorConfig.tabs[idx + 1], constructorConfig.tabs[idx]] = [constructorConfig.tabs[idx], constructorConfig.tabs[idx + 1]];
                    } else if (act === "del") {
                        if (!confirm(`Удалить вкладку «${constructorConfig.tabs[idx].label}»?`)) return;
                        constructorConfig.tabs.splice(idx, 1);
                    }
                    persistConstructorConfig();
                    renderConstructorTabsList();
                    applyConstructorConfig();
                });
            });
        }

        function addScheduleTab() {
            const label = document.getElementById("constrNewTabLabel").value.trim();
            const icon = document.getElementById("constrNewTabIcon").value.trim() || "📄";
            const content = document.getElementById("constrNewTabContent").value.trim();
            if (!label) { setConstrStatus("❌ Укажите название вкладки", "error"); return; }

            const key = "custom_" + Date.now();
            const tabId = "custom_" + Date.now();

            constructorConfig.tabs.push({
                key: key,
                label: label,
                icon: icon,
                visible: true,
                content: content
            });

            if (content) {
                constructorConfig.customTabs.push({
                    id: tabId,
                    label: label,
                    icon: icon,
                    content: content
                });
            }

            persistConstructorConfig();
            renderConstructorTabsList();
            renderConstructorCustomTabsList();
            applyConstructorConfig();

            document.getElementById("constrNewTabLabel").value = "";
            document.getElementById("constrNewTabIcon").value = "";
            document.getElementById("constrNewTabContent").value = "";

            setConstrStatus(`✅ Вкладка «${label}» добавлена`, "success");
        }

        function renderConstructorCustomTabsList() {
            const container = document.getElementById("constrCustomTabsList");
            if (!container) return;
            container.innerHTML = "";

            if (constructorConfig.customTabs.length === 0) {
                container.innerHTML = '<p class="constr-hint">Пока нет пользовательских вкладок. Добавьте первую ниже.</p>';
                return;
            }

            constructorConfig.customTabs.forEach((tab, idx) => {
                const row = document.createElement("div");
                row.className = "constr-tab-row";
                row.style.gridTemplateColumns = "40px 1fr 80px 120px";
                row.innerHTML = `
                    <div class="constr-tab-idx">${idx + 1}</div>
                    <input type="text" value="${escapeHtml(tab.label)}" data-field="label" data-idx="${idx}" placeholder="Название">
                    <input type="text" value="${escapeHtml(tab.icon)}" data-field="icon" data-idx="${idx}" maxlength="4">
                    <div class="constr-tab-actions">
                        <button class="icon-btn" data-act="edit" data-idx="${idx}" title="Редактировать содержимое">✎</button>
                        <button class="icon-btn danger" data-act="del" data-idx="${idx}" title="Удалить">🗑</button>
                    </div>
                `;
                container.appendChild(row);
            });

            container.querySelectorAll("input[data-field]").forEach(inp => {
                inp.addEventListener("change", () => {
                    const idx = parseInt(inp.dataset.idx);
                    const field = inp.dataset.field;
                    constructorConfig.customTabs[idx][field] = inp.value;
                    persistConstructorConfig();
                    applyConstructorConfig();
                });
            });

            container.querySelectorAll("[data-act]").forEach(btn => {
                btn.addEventListener("click", () => {
                    const idx = parseInt(btn.dataset.idx);
                    if (btn.dataset.act === "del") {
                        if (!confirm(`Удалить «${constructorConfig.customTabs[idx].label}»?`)) return;
                        constructorConfig.customTabs.splice(idx, 1);
                        persistConstructorConfig();
                        renderConstructorCustomTabsList();
                        applyConstructorConfig();
                    } else if (btn.dataset.act === "edit") {
                        const tab = constructorConfig.customTabs[idx];
                        const newContent = prompt("Введите HTML-содержимое:", tab.content);
                        if (newContent !== null) {
                            tab.content = newContent;
                            persistConstructorConfig();
                            applyConstructorConfig();
                            setConstrStatus(`✓ Содержимое «${tab.label}» обновлено`, "success");
                        }
                    }
                });
            });
        }

        function addCustomTab() {
            const label = document.getElementById("constrCustomTabLabel").value.trim();
            const icon = document.getElementById("constrCustomTabIcon").value.trim() || "📄";
            const content = document.getElementById("constrCustomTabContent").value.trim();
            if (!label) { setConstrStatus("❌ Укажите название", "error"); return; }
            if (!content) { setConstrStatus("❌ Добавьте HTML-содержимое", "error"); return; }

            const id = "custom_" + Date.now();
            constructorConfig.customTabs.push({ id, label, icon, content });

            constructorConfig.tabs.push({
                key: "custom:" + id,
                label: label,
                icon: icon,
                visible: true,
                isCustom: true,
                customId: id
            });

            persistConstructorConfig();
            renderConstructorTabsList();
            renderConstructorCustomTabsList();
            applyConstructorConfig();
            bindCustomTabClicks();

            document.getElementById("constrCustomTabLabel").value = "";
            document.getElementById("constrCustomTabIcon").value = "";
            document.getElementById("constrCustomTabContent").value = "";

            setConstrStatus(`✅ Вкладка «${label}» создана`, "success");
        }

        function applyThemePreset(name) {
            const presets = {
                classic: {
                    accent: "#1a3a6b", bgPage: "#eef3f7", bgContainer: "#ffffff",
                    textPrimary: "#0b2a4a", textSecondary: "#3d5a7a", border: "#d0ddee",
                    substBg: "#fff3cd", substText: "#7a5500", success: "#2a6b3a", danger: "#b00020"
                },
                dark: {
                    accent: "#4a82c7", bgPage: "#0d1520", bgContainer: "#16202f",
                    textPrimary: "#e6eef8", textSecondary: "#a8bdd4", border: "#2a3c56",
                    substBg: "#3a2f10", substText: "#f0c040", success: "#3d8f52", danger: "#e05070"
                },
                sea: {
                    accent: "#0f5d8c", bgPage: "#e6f4f9", bgContainer: "#ffffff",
                    textPrimary: "#0b2a4a", textSecondary: "#3d5a7a", border: "#a8d3e6",
                    substBg: "#fff3cd", substText: "#7a5500", success: "#1a7a5e", danger: "#b00020"
                },
                forest: {
                    accent: "#2a5a3a", bgPage: "#eef5ef", bgContainer: "#ffffff",
                    textPrimary: "#0b2a1a", textSecondary: "#3d5a4a", border: "#c0ddc8",
                    substBg: "#f5f3cc", substText: "#5a5500", success: "#2a6b3a", danger: "#8a1a20"
                },
                sunset: {
                    accent: "#8b3a1a", bgPage: "#fdf3ec", bgContainer: "#ffffff",
                    textPrimary: "#3a1a0b", textSecondary: "#7a4a3a", border: "#e6c8b8",
                    substBg: "#fff3cd", substText: "#7a5500", success: "#7a5a2a", danger: "#b00020"
                },
                minimal: {
                    accent: "#333333", bgPage: "#f5f5f5", bgContainer: "#ffffff",
                    textPrimary: "#1a1a1a", textSecondary: "#555555", border: "#e0e0e0",
                    substBg: "#f0f0f0", substText: "#333333", success: "#2a7a3a", danger: "#b00020"
                },
                purple: {
                    accent: "#5a2a8a", bgPage: "#f3eefa", bgContainer: "#ffffff",
                    textPrimary: "#2a0b4a", textSecondary: "#5a3a7a", border: "#d5c8e6",
                    substBg: "#fff3cd", substText: "#5a3a00", success: "#2a6b3a", danger: "#b00020"
                }
            };

            const preset = presets[name];
            if (!preset) return;
            Object.assign(constructorConfig.theme, preset);
            persistConstructorConfig();
            fillConstructorFields();
            applyConstructorConfig();
            setConstrStatus(`🎨 Пресет «${name}» применён`, "success");
        }

        function exportConstructorConfig() {
            const data = {
                exportedAt: new Date().toISOString(),
                type: "constructor-config",
                version: 1,
                config: constructorConfig
            };
            const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = `constructor_${formatDateForFile()}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            setConstrStatus("📤 Конфигурация скачана", "success");
        }

        function importConstructorConfig(e) {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (ev) => {
                try {
                    const parsed = JSON.parse(ev.target.result);
                    const cfg = parsed.config || parsed;
                    if (!cfg || typeof cfg !== "object") throw new Error("Неверный формат");
                    constructorConfig = deepMerge(JSON.parse(JSON.stringify(FACTORY_CONFIG)), cfg);
                    persistConstructorConfig();
                    fillConstructorFields();
                    renderConstructorTabsList();
                    renderConstructorCustomTabsList();
                    applyConstructorConfig();
                    setConstrStatus("✅ Конфигурация импортирована", "success");
                } catch (err) {
                    setConstrStatus(`❌ Ошибка: ${err.message}`, "error");
                }
                e.target.value = "";
            };
            reader.readAsText(file, "utf-8");
        }

        function initConstructorTab() {
            constructorConfig = loadConstructorConfig();

            const navBtns = document.querySelectorAll(".constr-nav-btn");
            navBtns.forEach(btn => {
                btn.addEventListener("click", () => {
                    navBtns.forEach(b => b.classList.remove("active"));
                    btn.classList.add("active");
                    document.querySelectorAll(".constr-section").forEach(s => s.classList.remove("active"));
                    const sec = document.querySelector(`.constr-section[data-section="${btn.dataset.section}"]`);
                    if (sec) sec.classList.add("active");
                });
            });

            fillConstructorFields();
            bindConstructorInputs();
            renderConstructorTabsList();
            renderConstructorCustomTabsList();

            document.getElementById("constrSaveBtn").addEventListener("click", () => {
                collectConstructorFields();
                persistConstructorConfig();
                applyConstructorConfig();
                setConstrStatus("✅ Настройки применены и сохранены", "success");
            });
            document.getElementById("constrExportBtn").addEventListener("click", exportConstructorConfig);
            document.getElementById("constrImportBtn").addEventListener("click", () => {
                document.getElementById("constrImportFile").click();
            });
            document.getElementById("constrImportFile").addEventListener("change", importConstructorConfig);
            document.getElementById("constrResetBtn").addEventListener("click", () => {
                if (!confirm("Сбросить все настройки конструктора к последнему сохранённому состоянию?")) return;
                constructorConfig = loadConstructorConfig();
                fillConstructorFields();
                renderConstructorTabsList();
                renderConstructorCustomTabsList();
                applyConstructorConfig();
                setConstrStatus("↺ Сброшено к сохранённому", "success");
            });
            document.getElementById("constrFactoryBtn").addEventListener("click", () => {
                if (!confirm("⚠ Сбросить ВСЁ к заводским настройкам? Текущие изменения будут потеряны.")) return;
                constructorConfig = JSON.parse(JSON.stringify(FACTORY_CONFIG));
                persistConstructorConfig();
                fillConstructorFields();
                renderConstructorTabsList();
                renderConstructorCustomTabsList();
                applyConstructorConfig();
                setConstrStatus("🏭 Заводские настройки восстановлены", "success");
            });

            document.querySelectorAll(".constr-preset-btn").forEach(btn => {
                btn.addEventListener("click", () => applyThemePreset(btn.dataset.preset));
            });

            document.getElementById("constrAddTabBtn").addEventListener("click", addScheduleTab);
            document.getElementById("constrAddCustomTabBtn").addEventListener("click", addCustomTab);
        }

        // ==================== ОБНОВЛЕНИЕ САЙТА ====================
        function getCurrentHtmlSource() {
            const doctype = "<!DOCTYPE html>\n";
            const html = document.documentElement.outerHTML;
            return doctype + html;
        }

        function setUpdateStatus(msg, type) {
            const el = document.getElementById("updateStatus");
            if (!el) return;
            el.textContent = msg;
            el.className = "update-status" + (type ? " " + type : "");
            clearTimeout(el._timer);
            el._timer = setTimeout(() => {
                el.textContent = "";
                el.className = "update-status";
            }, 5000);
        }

        function formatBytes(bytes) {
            if (bytes === 0) return "0 Б";
            if (bytes < 1024) return bytes + " Б";
            if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " КБ";
            return (bytes / 1024 / 1024).toFixed(2) + " МБ";
        }

        function refreshUpdateInfo() {
            try {
                const html = getCurrentHtmlSource();
                const sizeEl = document.getElementById("updateCurrentSize");
                if (sizeEl) sizeEl.textContent = formatBytes(new Blob([html]).size);
            } catch (e) {
                const sizeEl = document.getElementById("updateCurrentSize");
                if (sizeEl) sizeEl.textContent = "—";
            }

            const versionMeta = document.querySelector('meta[name="app-version"]');
            const versionEl = document.getElementById("updateCurrentVersion");
            if (versionEl) versionEl.textContent = versionMeta ? versionMeta.content : "не указана";

            try {
                const last = localStorage.getItem(STORAGE_KEYS.lastUpdate);
                const lastEl = document.getElementById("updateLastUpdate");
                if (lastEl) {
                    if (last) {
                        const d = new Date(last);
                        lastEl.textContent = d.toLocaleString("ru-RU");
                    } else {
                        lastEl.textContent = "не производилось";
                    }
                }
            } catch (e) {}

            let totalDataSize = 0;
            Object.values(STORAGE_KEYS).forEach(k => {
                if (k === "updateBackup" || k === "updateDataBackup") return;
                try {
                    totalDataSize += (localStorage.getItem(k) || "").length;
                } catch (e) {}
            });
            const dataSizeEl = document.getElementById("updateDataSize");
            if (dataSizeEl) dataSizeEl.textContent = formatBytes(totalDataSize);

            try {
                const backup = localStorage.getItem(STORAGE_KEYS.updateBackup);
                const backupStatusEl = document.getElementById("updateBackupStatus");
                const rollbackBtn = document.getElementById("updateRollbackBtn");
                const clearBackupBtn = document.getElementById("updateClearBackupBtn");
                if (backup) {
                    const meta = JSON.parse(localStorage.getItem(STORAGE_KEYS.updateBackup + "_meta") || "{}");
                    const date = meta.savedAt ? new Date(meta.savedAt).toLocaleString("ru-RU") : "—";
                    if (backupStatusEl) backupStatusEl.textContent = `сохранена (${date})`;
                    if (rollbackBtn) rollbackBtn.disabled = false;
                    if (clearBackupBtn) clearBackupBtn.disabled = false;
                } else {
                    if (backupStatusEl) backupStatusEl.textContent = "не найдена";
                    if (rollbackBtn) rollbackBtn.disabled = true;
                    if (clearBackupBtn) clearBackupBtn.disabled = true;
                }
            } catch (e) {}

            try {
                const dataBackup = localStorage.getItem(STORAGE_KEYS.updateDataBackup);
                const dataBackupStatusEl = document.getElementById("updateDataBackupStatus");
                const restoreBtn = document.getElementById("updateRestoreDataBtn");
                if (dataBackup) {
                    const meta = JSON.parse(dataBackup);
                    const date = meta.savedAt ? new Date(meta.savedAt).toLocaleString("ru-RU") : "—";
                    if (dataBackupStatusEl) dataBackupStatusEl.textContent = date;
                    if (restoreBtn) restoreBtn.disabled = false;
                } else {
                    if (dataBackupStatusEl) dataBackupStatusEl.textContent = "не производился";
                    if (restoreBtn) restoreBtn.disabled = true;
                }
            } catch (e) {}
        }

        function downloadCurrentHtml() {
            try {
                const html = getCurrentHtmlSource();
                const blob = new Blob([html], { type: "text/html;charset=utf-8" });
                const url = URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url;
                a.download = `raspisanie_${formatDateForFile()}.html`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                setUpdateStatus("✅ HTML-файл скачан", "success");
            } catch (err) {
                console.error(err);
                setUpdateStatus("❌ Не удалось получить исходный код: " + err.message, "error");
            }
        }

        function copyCurrentHtml() {
            try {
                const html = getCurrentHtmlSource();
                navigator.clipboard.writeText(html)
                    .then(() => {
                        setUpdateStatus("📋 HTML скопирован в буфер обмена", "success");
                    })
                    .catch(() => {
                        setUpdateStatus("❌ Не удалось скопировать в буфер", "error");
                    });
            } catch (err) {
                console.error(err);
                setUpdateStatus("❌ Не удалось получить исходный код: " + err.message, "error");
            }
        }

        function backupAllData() {
            const data = {};
            Object.values(STORAGE_KEYS).forEach(k => {
                if (k === "updateBackup" || k === "updateDataBackup") return;
                try {
                    const v = localStorage.getItem(k);
                    if (v) data[k] = v;
                } catch (e) {}
            });

            const backup = {
                savedAt: new Date().toISOString(),
                version: 1,
                data: data
            };

            try {
                localStorage.setItem(STORAGE_KEYS.updateDataBackup, JSON.stringify(backup));
                setUpdateStatus("💾 Резервная копия данных создана", "success");
                refreshUpdateInfo();
            } catch (e) {
                setUpdateStatus("❌ Ошибка создания бэкапа: " + e.message, "error");
                return;
            }

            const blob = new Blob([JSON.stringify(backup, null, 2)], { type: "application/json" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = `backup_data_${formatDateForFile()}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        function restoreDataFromBackup() {
            try {
                const raw = localStorage.getItem(STORAGE_KEYS.updateDataBackup);
                if (!raw) { setUpdateStatus("❌ Бэкап данных не найден", "error"); return; }
                const backup = JSON.parse(raw);
                if (!backup.data) { setUpdateStatus("❌ Некорректный бэкап", "error"); return; }
                if (!confirm("⚠ Восстановить все данные из бэкапа от " + new Date(backup.savedAt).toLocaleString("ru-RU") + "?\n\nТекущие данные будут заменены.")) return;

                Object.entries(backup.data).forEach(([k, v]) => {
                    try {
                        localStorage.setItem(k, v);
                    } catch (e) {}
                });

                setUpdateStatus("✅ Данные восстановлены. Перезагрузка...", "success");
                setTimeout(() => location.reload(), 1200);
            } catch (e) {
                setUpdateStatus("❌ Ошибка восстановления: " + e.message, "error");
            }
        }

        function parseHtmlVersion(html) {
            const metaMatch = html.match(/<meta\s+name=["']app-version["']\s+content=["']([^"']+)["']/i);
            if (metaMatch) return metaMatch[1];
            const commentMatch = html.match(/<!--\s*app-version:\s*([^\s-]+)\s*-->/i);
            if (commentMatch) return commentMatch[1];
            return null;
        }

        function validateHtmlFile(html) {
            return {
                doctype: /<!DOCTYPE html>/i.test(html),
                html: /<html[\s>]/i.test(html),
                schedule: html.includes('id="scheduleTable"') || html.includes('id="scheduleBody"'),
                console: html.includes('id="consoleModal"') || html.includes('console-tab'),
                script: /<script[\s>]/i.test(html)
            };
        }

        function renderValidationChecks(checks) {
            const container = document.getElementById("updateChecks");
            if (!container) return;
            container.style.display = "block";

            const labels = {
                doctype: "DOCTYPE html",
                html: "Тег <html>",
                schedule: "Наличие расписания (scheduleTable/scheduleBody)",
                console: "Наличие консоли (consoleModal)",
                script: "Наличие <script>-блоков"
            };

            container.querySelectorAll(".update-check-item").forEach(item => {
                const key = item.dataset.check;
                const ok = checks[key];
                item.classList.toggle("ok", ok);
                item.classList.toggle("fail", !ok);
                item.querySelector(".update-check-icon").textContent = ok ? "✅" : "❌";
                item.querySelector(".update-check-text").textContent = labels[key] || key;
            });
        }

        function handleHtmlFile(file) {
            if (!file) return;
            if (!/\.html?$/i.test(file.name)) {
                setUpdateStatus("❌ Только файлы .html или .htm", "error");
                return;
            }

            const errEl = document.getElementById("updateError");
            if (errEl) errEl.textContent = "";

            if (file.size > 5 * 1024 * 1024) {
                setUpdateStatus("❌ Файл слишком большой (макс. 5 МБ)", "error");
                return;
            }

            const reader = new FileReader();
            reader.onload = (ev) => {
                const html = ev.target.result;

                const checks = validateHtmlFile(html);
                renderValidationChecks(checks);

                const allOk = Object.values(checks).every(v => v);
                const criticalOk = checks.doctype && checks.html && checks.script;

                pendingUpdateHtml = html;
                pendingUpdateMeta = {
                    name: file.name,
                    size: file.size,
                    date: new Date(file.lastModified),
                    version: parseHtmlVersion(html),
                    allChecksOk: allOk,
                    criticalOk: criticalOk
                };

                document.getElementById("updateFileInfo").style.display = "block";
                document.getElementById("updateFileName").textContent = file.name;
                document.getElementById("updateFileSize").textContent = formatBytes(file.size);
                document.getElementById("updateFileDate").textContent = new Date(file.lastModified).toLocaleString("ru-RU");
                document.getElementById("updateFileVersion").textContent = pendingUpdateMeta.version || "не указана";

                document.getElementById("updateActions").style.display = "flex";

                const applyBtn = document.getElementById("updateApplyBtn");
                if (criticalOk) {
                    applyBtn.disabled = false;
                    if (allOk) {
                        setUpdateStatus("✅ Файл прошёл все проверки. Можно применять.", "success");
                    } else {
                        setUpdateStatus("⚠ Некоторые проверки не пройдены, но критические — OK.", "warning");
                    }
                } else {
                    applyBtn.disabled = true;
                    setUpdateStatus("❌ Файл не похож на приложение «Расписание»", "error");
                    if (errEl) errEl.textContent = "❌ Отсутствуют критичные элементы: DOCTYPE, <html> или <script>";
                }

                const dz = document.getElementById("updateDropZone");
                if (dz) dz.classList.add("has-file");
            };
            reader.readAsText(file, "utf-8");
        }

        function applyUpdate() {
            if (!pendingUpdateHtml) {
                setUpdateStatus("❌ Нет загруженного файла", "error");
                return;
            }

            if (!confirm("⚠ Применить обновление?\n\n1. Все текущие данные будут сохранены.\n2. Новый HTML заменит текущий.\n3. Страница перезагрузится.\n\nПродолжить?")) return;

            try {
                backupAllData();
            } catch (e) {
                console.error("Ошибка при создании бэкапа данных:", e);
            }

            let currentHtml = "";
            try {
                currentHtml = getCurrentHtmlSource();
            } catch (e) {
                console.error("Не удалось получить текущий HTML:", e);
            }

            if (currentHtml) {
                try {
                    localStorage.setItem(STORAGE_KEYS.updateBackup, currentHtml);
                    localStorage.setItem(STORAGE_KEYS.updateBackup + "_meta", JSON.stringify({
                        savedAt: new Date().toISOString(),
                        version: document.querySelector('meta[name="app-version"]')?.content || null,
                        size: currentHtml.length
                    }));
                } catch (e) {
                    console.error("Не удалось сохранить бэкап HTML:", e);
                    if (e.name === "QuotaExceededError") {
                        if (!confirm("⚠ Не удалось сохранить резервную копию HTML (превышен лимит хранилища).\n\nПродолжить без отката?")) return;
                    }
                }
            }

            try {
                localStorage.setItem(STORAGE_KEYS.lastUpdate, new Date().toISOString());
            } catch (e) {}

            try {
                document.open();
                document.write(pendingUpdateHtml);
                document.close();

                setTimeout(() => {
                    location.reload();
                }, 100);
            } catch (e) {
                setUpdateStatus("❌ Ошибка применения: " + e.message, "error");
            }
        }

        function cancelUpdate() {
            pendingUpdateHtml = null;
            pendingUpdateMeta = null;
            document.getElementById("updateFileInfo").style.display = "none";
            document.getElementById("updateChecks").style.display = "none";
            document.getElementById("updateActions").style.display = "none";
            document.getElementById("updateError").textContent = "";
            document.getElementById("updateHtmlFile").value = "";
            const dz = document.getElementById("updateDropZone");
            if (dz) dz.classList.remove("has-file");
            setUpdateStatus("Отменено");
        }

        function rollbackUpdate() {
            try {
                const backup = localStorage.getItem(STORAGE_KEYS.updateBackup);
                if (!backup) { setUpdateStatus("❌ Резервная копия не найдена", "error"); return; }
                const meta = JSON.parse(localStorage.getItem(STORAGE_KEYS.updateBackup + "_meta") || "{}");
                const date = meta.savedAt ? new Date(meta.savedAt).toLocaleString("ru-RU") : "—";
                if (!confirm(`⏪ Откатить к версии от ${date}?\n\nТекущий HTML будет заменён резервной копией.`)) return;

                document.open();
                document.write(backup);
                document.close();
                setTimeout(() => location.reload(), 100);
            } catch (e) {
                setUpdateStatus("❌ Ошибка отката: " + e.message, "error");
            }
        }

        function clearHtmlBackup() {
            if (!confirm("🗑 Удалить резервную копию HTML?")) return;
            try {
                localStorage.removeItem(STORAGE_KEYS.updateBackup);
                localStorage.removeItem(STORAGE_KEYS.updateBackup + "_meta");
                setUpdateStatus("🗑 Резервная копия HTML удалена", "success");
                refreshUpdateInfo();
            } catch (e) {
                setUpdateStatus("❌ Ошибка: " + e.message, "error");
            }
        }

        function initUpdateTab() {
            document.getElementById("updateDownloadHtmlBtn").addEventListener("click", downloadCurrentHtml);
            document.getElementById("updateCopyHtmlBtn").addEventListener("click", copyCurrentHtml);
            document.getElementById("updateBackupDataBtn").addEventListener("click", backupAllData);

            document.getElementById("updateSelectFileBtn").addEventListener("click", (e) => {
                e.stopPropagation();
                document.getElementById("updateHtmlFile").click();
            });
            document.getElementById("updateDropZone").addEventListener("click", () => {
                document.getElementById("updateHtmlFile").click();
            });
            document.getElementById("updateHtmlFile").addEventListener("change", (e) => {
                const file = e.target.files[0];
                handleHtmlFile(file);
            });

            const dz = document.getElementById("updateDropZone");
            ["dragenter", "dragover"].forEach(evt => {
                dz.addEventListener(evt, (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    dz.classList.add("dragover");
                });
            });
            ["dragleave", "drop"].forEach(evt => {
                dz.addEventListener(evt, (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    dz.classList.remove("dragover");
                });
            });
            dz.addEventListener("drop", (e) => {
                const file = e.dataTransfer.files[0];
                if (file) handleHtmlFile(file);
            });

            document.getElementById("updateApplyBtn").addEventListener("click", applyUpdate);
            document.getElementById("updateCancelBtn").addEventListener("click", cancelUpdate);

            document.getElementById("updateRollbackBtn").addEventListener("click", rollbackUpdate);
            document.getElementById("updateClearBackupBtn").addEventListener("click", clearHtmlBackup);

            document.getElementById("updateRestoreDataBtn").addEventListener("click", restoreDataFromBackup);

            refreshUpdateInfo();
        }

        // ==================== WORD-ИНСТРУКЦИЯ ====================
        const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
                Table, TableRow, TableCell, WidthType, BorderStyle, PageBreak } = docx;
        const SCHOOL_NAME = "МОУ «Северная СОШ №2»";
        const SCHOOL_SUBTITLE = "Белгородского муниципального округа Белгородской области";

        function createTextParagraph(text, opts = {}) {
            return new Paragraph({
                children: [
                    new TextRun({
                        text: text,
                        bold: opts.bold || false,
                        italics: opts.italics || false,
                        size: opts.size ? opts.size * 2 : 22,
                        font: "Times New Roman",
                        color: opts.color || "1F2A44"
                    })
                ],
                alignment: opts.align || AlignmentType.JUSTIFIED,
                spacing: { before: opts.spaceBefore || 100, after: opts.spaceAfter || 100, line: 320 },
                heading: opts.heading || undefined,
                indent: opts.indent || undefined
            });
        }

        function createHeading1(text) {
            return new Paragraph({
                children: [new TextRun({ text: text, bold: true, size: 32, font: "Times New Roman", color: "0B2A4A" })],
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 400, after: 200, line: 360 }
            });
        }

        function createHeading2(text) {
            return new Paragraph({
                children: [new TextRun({ text: text, bold: true, size: 26, font: "Times New Roman", color: "1A3A6B" })],
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 300, after: 150, line: 340 }
            });
        }

        function createBullet(text) {
            return new Paragraph({
                children: [new TextRun({ text: text, size: 22, font: "Times New Roman", color: "1F2A44" })],
                bullet: { level: 0 },
                spacing: { before: 40, after: 40, line: 320 }
            });
        }

        function createNumberedItem(text, num = 1) {
            return new Paragraph({
                children: [new TextRun({ text: `${num}. ${text}`, size: 22, font: "Times New Roman", color: "1F2A44" })],
                spacing: { before: 40, after: 40, line: 320 },
                indent: { left: 360 }
            });
        }

        function createTable(headers, rows, colWidths) {
            const tableRows = [];
            tableRows.push(new TableRow({
                tableHeader: true,
                children: headers.map((h, i) => new TableCell({
                    children: [
                        new Paragraph({
                            children: [new TextRun({ text: h, bold: true, size: 22, font: "Times New Roman", color: "FFFFFF" })],
                            alignment: AlignmentType.CENTER,
                            spacing: { before: 60, after: 60 }
                        })
                    ],
                    shading: { fill: "1A3A6B" },
                    width: colWidths ? { size: colWidths[i], type: WidthType.PERCENTAGE } : undefined,
                    margins: { top: 80, bottom: 80, left: 100, right: 100 }
                }))
            }));

            rows.forEach((row, rowIdx) => {
                tableRows.push(new TableRow({
                    children: row.map((cell, i) => new TableCell({
                        children: [
                            new Paragraph({
                                children: [new TextRun({ text: String(cell), size: 20, font: "Times New Roman", color: "1F2A44" })],
                                spacing: { before: 40, after: 40 }
                            })
                        ],
                        shading: { fill: rowIdx % 2 === 0 ? "F5F8FC" : "FFFFFF" },
                        width: colWidths ? { size: colWidths[i], type: WidthType.PERCENTAGE } : undefined,
                        margins: { top: 60, bottom: 60, left: 100, right: 100 }
                    }))
                }));
            });

            return new Table({
                rows: tableRows,
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.SINGLE, size: 4, color: "1A3A6B" },
                    bottom: { style: BorderStyle.SINGLE, size: 4, color: "1A3A6B" },
                    left: { style: BorderStyle.SINGLE, size: 4, color: "1A3A6B" },
                    right: { style: BorderStyle.SINGLE, size: 4, color: "1A3A6B" },
                    insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "B0C0D0" },
                    insideVertical: { style: BorderStyle.SINGLE, size: 2, color: "B0C0D0" }
                }
            });
        }

        async function generateAdminGuide() {
            const statusEl = document.getElementById("guideStatus");
            if (statusEl) statusEl.textContent = "⏳ Формирование...";
            try {
                const children = [];

                children.push(new Paragraph({ children: [], spacing: { before: 1200, after: 0 } }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: SCHOOL_NAME, bold: true, size: 48, font: "Times New Roman", color: "0B2A4A" })],
                    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 200 }
                }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: SCHOOL_SUBTITLE, size: 24, font: "Times New Roman", color: "3D5A7A" })],
                    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 1200 }
                }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: "ПОЛНАЯ ИНСТРУКЦИЯ ПОЛЬЗОВАТЕЛЯ", bold: true, size: 52, font: "Times New Roman", color: "1A3A6B" })],
                    alignment: AlignmentType.CENTER, spacing: { before: 600, after: 200 }
                }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: "Веб-приложение «Расписание учителей»", italics: true, size: 28, font: "Times New Roman", color: "3D5A7A" })],
                    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 1600 }
                }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: `Версия документа: 10.2`, size: 22, font: "Times New Roman", color: "7A8FA8" })],
                    alignment: AlignmentType.CENTER, spacing: { before: 200, after: 100 }
                }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: `Дата: ${new Date().toLocaleDateString("ru-RU")}`, size: 22, font: "Times New Roman", color: "7A8FA8" })],
                    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 100 }
                }));

                children.push(new Paragraph({ children: [new PageBreak()] }));

                children.push(createHeading1("1. Общие сведения о приложении"));
                children.push(createTextParagraph("«Расписание учителей» — это веб-приложение (один HTML-файл), которое работает прямо в браузере и не требует установки на сервер или в операционную систему. Все данные хранятся локально в браузере (localStorage), поэтому приложение работает без интернета."));
                children.push(createTextParagraph("Приложение предназначено для:", { bold: true }));
                children.push(createBullet("заместителей директора по учебной работе;"));
                children.push(createBullet("администраторов школы;"));
                children.push(createBullet("учителей, которым нужно посмотреть своё расписание."));

                children.push(createHeading2("1.1. Из чего состоит приложение"));
                children.push(createTextParagraph("Верхняя часть экрана — шапка с названием школы, вкладками расписания и основными кнопками. Ниже — панель поиска, фильтр по дням недели и сама таблица расписания. В самом низу — строка подсказки и статус."));
                children.push(createTextParagraph("Всего в приложении четыре вкладки расписания и одна консоль администратора (доступна по кнопке «⚙ Консоль»)."));

                children.push(new Paragraph({ children: [new PageBreak()] }));
                children.push(createHeading1("2. Вкладки расписания"));
                children.push(createTextParagraph("Переключение между вкладками происходит по клику на кнопку в шапке. Каждая вкладка показывает своё расписание."));
                children.push(createTable(
                    ["Вкладка", "Что показывает", "Кто может редактировать"],
                    [
                        ["👨‍🏫 Учительское", "Расписание по учителям: предметы, замены, кабинеты.", "Администратор"],
                        ["🎓 Детское", "Расписание по классам: какой урок, какой учитель.", "Администратор"],
                        ["📚 ИУП 5-8", "Индивидуальные учебные планы учеников с ОВЗ.", "Администратор"],
                        ["📌 Журнал замен", "Учёт замен для бухгалтерии: кто, когда и кого заменял.", "Администратор"]
                    ],
                    [22, 50, 28]
                ));

                children.push(createHeading2("2.1. Вкладка «👨‍🏫 Учительское»"));
                children.push(createTextParagraph("Это основное расписание. В левой колонке — имя учителя, затем кабинет и класс (классное руководство). Далее — дни недели (Пн–Пт) и по 8 уроков в каждом дне."));
                children.push(createTextParagraph("Что можно делать:", { bold: true }));
                children.push(createBullet("Искать учителя через поиск или выпадающий список."));
                children.push(createBullet("Фильтровать по дню недели (кнопки «Пн», «Вт» и т. д.)."));
                children.push(createBullet("Включить чекбокс «📌 Только замены» — останутся только ячейки с заменами."));
                children.push(createBullet("В режиме редактирования — кликнуть на ячейку и изменить предмет или замену."));
                children.push(createTextParagraph("Особенность:", { bold: true }));
                children.push(createTextParagraph("Когда вы вписываете замену, приложение показывает список свободных учителей на этот урок. Это позволяет не ошибиться и не поставить замену тому, кто сам ведёт урок."));

                children.push(createHeading2("2.2. Вкладка «🎓 Детское»"));
                children.push(createTextParagraph("Здесь расписание представлено по классам. В левой колонке — название класса, далее — кабинет и дни недели с уроками."));
                children.push(createTextParagraph("Что можно делать:", { bold: true }));
                children.push(createBullet("Искать класс (например, «5а»)."));
                children.push(createBullet("Редактировать предметы и указывать замены (ФИО вводится вручную)."));
                children.push(createBullet("Смотреть расписание класса в режиме «Моё расписание»."));

                children.push(createHeading2("2.3. Вкладка «📚 ИУП 5-8»"));
                children.push(createTextParagraph("ИУП — индивидуальный учебный план. Это расписание для отдельных учеников (например, детей с ОВЗ). В левой колонке — имя ученика, затем класс и уроки."));
                children.push(createTextParagraph("Что можно делать:", { bold: true }));
                children.push(createBullet("Искать ученика по имени."));
                children.push(createBullet("Редактировать занятия и указывать замены вручную."));
                children.push(createBullet("Печатать индивидуальное расписание ученика."));

                children.push(createHeading2("2.4. Вкладка «📌 Журнал замен»"));
                children.push(createTextParagraph("Это отдельный модуль для учёта замен. Он нужен для бухгалтерии и отчётности."));
                children.push(createTextParagraph("Разделы журнала:", { bold: true }));
                children.push(createBullet("Форма «Новая замена» — сверху. Указывается дата, тип урока (стандартный или внеурочная деятельность), отсутствующий учитель и его предмет, заменяющий учитель и его предмет."));
                children.push(createBullet("Дашборд — четыре карточки со статистикой: всего замен, сколько учителей заменено, сколько учителей заменяло, сколько замен за текущий месяц."));
                children.push(createBullet("Панель фильтров — переключатель «📅 Месяц / 📆 День», поиск, кнопки «🗓 Календарь», «⚙ Справочники», «📤 Excel», «🖨 Печать»."));
                children.push(createBullet("Таблица замен — список всех записей с кнопками «✎ Редактировать» и «🗑 Удалить»."));
                children.push(createBullet("Мини-календарь — включается кнопкой «🗓 Календарь». Дни с заменами подсвечиваются, по клику открывается нужная дата."));
                children.push(createBullet("История журнала — кнопка «📜 История журнала». Позволяет отменить последнее действие."));
                children.push(createTextParagraph("Экспорт и печать:", { bold: true }));
                children.push(createBullet("«📤 Excel» — скачивает журнал замен в формате Excel с колонкой «Тип урока»."));
                children.push(createBullet("«🖨 Печать» — формирует печатную форму с колонкой «Подпись»."));

                children.push(new Paragraph({ children: [new PageBreak()] }));
                children.push(createHeading1("3. Режим редактирования"));
                children.push(createTextParagraph("Чтобы редактировать расписание, нажмите кнопку «✎ Редактировать» в шапке. Появится окно ввода пароля. По умолчанию пароль — sever2."));
                children.push(createTextParagraph("После входа в режим редактирования:", { bold: true }));
                children.push(createBullet("Появляется индикатор «Режим редактирования»."));
                children.push(createBullet("Становятся видны кнопки «↺ Сбросить» и «💾 Сохранить»."));
                children.push(createBullet("Ячейки таблицы становятся кликабельными."));
                children.push(createTextParagraph("Как редактировать ячейку:", { bold: true }));
                children.push(createNumberedItem("Кликните по ячейке — откроется мини-редактор.", 1));
                children.push(createNumberedItem("Введите название предмета в верхнее поле.", 2));
                children.push(createNumberedItem("В поле «📌 Замена» введите ФИО учителя. В режиме «Учительское» появится список свободных учителей — можно выбрать из него.", 3));
                children.push(createNumberedItem("Нажмите Enter — сохранить, Esc — отменить.", 4));
                children.push(createTextParagraph("Кнопки:", { bold: true }));
                children.push(createBullet("«💾 Сохранить» — сохраняет все изменения в браузере."));
                children.push(createBullet("«↺ Сбросить» — возвращает исходное (заводское) расписание."));
                children.push(createBullet("«🔒 Заблокировать» — выходит из режима редактирования."));

                children.push(createHeading1("4. Кнопка «👤 Моё расписание»"));
                children.push(createTextParagraph("Эта кнопка открывает персональное расписание. Что именно показывается — зависит от текущей вкладки:"));
                children.push(createBullet("На вкладке «Учительское» — расписание выбранного учителя."));
                children.push(createBullet("На вкладке «Детское» — расписание выбранного класса."));
                children.push(createBullet("На вкладке «ИУП 5-8» — расписание выбранного ученика."));
                children.push(createTextParagraph("В окне «Моё расписание» можно:", { bold: true }));
                children.push(createBullet("Выбрать учителя / класс / ученика из списка."));
                children.push(createBullet("Выбрать день: «Сегодня», «Завтра», конкретный день недели или «Вся неделя»."));
                children.push(createBullet("Нажать «🖨 Распечатать» — распечатать расписание."));
                children.push(createBullet("Нажать «🔔 Звонки» — настроить время начала и конца каждого урока."));
                children.push(createTextParagraph("Если учитель кого-то заменяет, внизу появится блок «🔔 В эти уроки вас заменяют» с указанием предмета, класса и заменяющего."));

                children.push(new Paragraph({ children: [new PageBreak()] }));
                children.push(createHeading1("5. Консоль администратора"));
                children.push(createTextParagraph("Консоль открывается кнопкой «⚙ Консоль» в шапке. Пароль по умолчанию — sever2. Консоль — это центр управления приложением."));
                children.push(createTable(
                    ["Вкладка", "Что делает", "Защита паролем"],
                    [
                        ["👥 Учителя", "Список учителей: добавление, удаление, редактирование, перемещение.", "Нет"],
                        ["⚡ Массовые операции", "Массовая очистка расписания и статистика загруженности.", "Нет"],
                        ["💾 Импорт/Экспорт", "Excel, JSON, инструкции Word.", "Нет"],
                        ["📜 История", "Полная история изменений с возможностью отмены.", "Нет"],
                        ["🔐 Безопасность", "Смена пароля администратора.", "Да"],
                        ["🎨 Конструктор", "Изменение внешнего вида приложения.", "Да"],
                        ["🔄 Обновление сайта", "Загрузка новой версии HTML-файла.", "Да"],
                        ["⚠ Опасная зона", "Полный сброс всех данных.", "Да"]
                    ],
                    [22, 60, 18]
                ));

                children.push(createHeading2("5.1. Вкладка «👥 Учителя»"));
                children.push(createBullet("Кнопка «+ Добавить учителя» — открывает окно с полями «Имя», «Кабинет», «Класс»."));
                children.push(createBullet("Поля в таблице можно редактировать прямо на месте."));
                children.push(createBullet("Кнопки «▲» и «▼» — меняют порядок учителя."));
                children.push(createBullet("Кнопка «🧹» — очищает расписание учителя."));
                children.push(createBullet("Кнопка «🗑» — удаляет учителя (с возможностью отмены через историю)."));

                children.push(createHeading2("5.2. Вкладка «⚡ Массовые операции»"));
                children.push(createTextParagraph("Слева — панель массового действия. Выберите одно из действий:", { bold: true }));
                children.push(createBullet("Очистить расписание учителя — выбрать учителя из списка."));
                children.push(createBullet("Очистить день у всех — выбрать день недели."));
                children.push(createBullet("Убрать все замены."));
                children.push(createBullet("Очистить всё расписание."));
                children.push(createTextParagraph("Справа — статистика загруженности: всего уроков, замен, учителей и разбивка по дням."));

                children.push(createHeading2("5.3. Вкладка «💾 Импорт/Экспорт»"));
                children.push(createTextParagraph("Excel (.xlsx / .xls):", { bold: true }));
                children.push(createBullet("«📥 Импорт» — загрузить расписание из Excel. Можно выбрать лист и строку-заголовок, посмотреть предпросмотр."));
                children.push(createBullet("«📤 Расписание» — скачать все три расписания (Учительское, Детское, ИУП) в один Excel-файл."));
                children.push(createBullet("«📜 История» — скачать историю изменений в Excel."));
                children.push(createTextParagraph("JSON (резервная копия):", { bold: true }));
                children.push(createBullet("«💾 Расписание» — полная копия всех данных в JSON."));
                children.push(createBullet("«📜 История» — копия истории."));
                children.push(createBullet("«📂 Импорт» — загрузить JSON-файл."));
                children.push(createBullet("«📋 Копировать» — скопировать JSON в буфер обмена."));
                children.push(createTextParagraph("Документация:", { bold: true }));
                children.push(createBullet("«📘 Полная» — скачать полную инструкцию (этот документ)."));
                children.push(createBullet("«📄 Краткая» — скачать краткую инструкцию."));

                children.push(createHeading2("5.4. Вкладка «📜 История»"));
                children.push(createTextParagraph("Показывает все изменения в приложении: кто, когда и что менял. Каждое изменение можно отменить кнопкой «↶». Есть кнопка «🗑 Очистить историю»."));

                children.push(createHeading2("5.5. Вкладка «🔐 Безопасность»"));
                children.push(createTextParagraph("Позволяет сменить пароль администратора. Введите текущий пароль, новый пароль (минимум 4 символа) и подтверждение. Есть индикатор надёжности пароля."));
                children.push(createTextParagraph("Также есть кнопка «🔓 Сбросить к sever2» — сбрасывает пароль без ввода текущего (требуется подтверждение)."));

                children.push(createHeading2("5.6. Вкладка «🎨 Конструктор»"));
                children.push(createTextParagraph("Конструктор позволяет менять внешний вид приложения без программирования. Разделы:"));
                children.push(createBullet("🏷️ Шапка — заголовок, подзаголовок, третья строка."));
                children.push(createBullet("📑 Вкладки — переименование, скрытие, изменение порядка."));
                children.push(createBullet("🎨 Цвета — настройка цветов и пресеты (Классика, Тёмная, Морская, Лесная, Закат, Минимализм, Фиолетовая)."));
                children.push(createBullet("📐 Размеры — шрифт, скругление, ширина контейнера, отступы."));
                children.push(createBullet("✨ Эффекты — анимированный фон, тени, blur."));
                children.push(createBullet("🔘 Кнопки — скрытие/показ кнопок в шапке (включая «Мобильная версия»)."));
                children.push(createBullet("📌 Футер — текст подсказки."));
                children.push(createBullet("🧩 Свои вкладки — создание собственных вкладок с HTML-содержимым."));
                children.push(createTextParagraph("Кнопки сверху: «💾 Применить», «📤 Экспорт», «📥 Импорт», «↺ Сброс», «🏭 Заводские»."));

                children.push(createHeading2("5.7. Вкладка «🔄 Обновление сайта»"));
                children.push(createTextParagraph("Эта вкладка нужна, когда вышла новая версия HTML-файла и нужно её установить, не потеряв данные."));
                children.push(createBullet("«📤 Скачать HTML» — сохранить текущий файл приложения."));
                children.push(createBullet("«📋 Копировать HTML» — скопировать код в буфер обмена."));
                children.push(createBullet("«💾 Бэкап данных» — создать резервную копию всех данных (JSON)."));
                children.push(createTextParagraph("Загрузка обновления:", { bold: true }));
                children.push(createNumberedItem("Перетащите новый HTML-файл в зону загрузки или нажмите «выберите файл».", 1));
                children.push(createNumberedItem("Приложение проверит файл: DOCTYPE, тег <html>, наличие расписания, консоли и script-блоков.", 2));
                children.push(createNumberedItem("Если проверки пройдены — нажмите «✅ Применить обновление».", 3));
                children.push(createNumberedItem("Перед обновлением автоматически создаётся бэкап данных и HTML.", 4));
                children.push(createTextParagraph("Откат:", { bold: true }));
                children.push(createBullet("«⏪ Откатить» — вернуть предыдущую версию HTML."));
                children.push(createBullet("«🗑 Удалить бэкап» — удалить резервную копию."));
                children.push(createBullet("«📂 Восстановить из бэкапа» — восстановить данные из JSON-бэкапа."));

                children.push(createHeading2("5.8. Вкладка «⚠ Опасная зона»"));
                children.push(createTextParagraph("Здесь находятся необратимые действия:"));
                children.push(createBullet("«🗑 Полный сброс» — сбросить все три расписания к исходным данным."));
                children.push(createBullet("«🧹 Очистить хранилище» — полностью очистить localStorage (все данные будут потеряны)."));
                children.push(createTextParagraph("Справа — информация о размере хранилища."));

                children.push(new Paragraph({ children: [new PageBreak()] }));
                children.push(createHeading1("6. Мобильная версия"));
                children.push(createTextParagraph("Кнопка «📱 Мобильная» в шапке переключает приложение в режим, оптимизированный для просмотра с телефона. Это полноценная переработка интерфейса, а не простое сжатие."));
                children.push(createTextParagraph("Что меняется в мобильном режиме:", { bold: true }));
                children.push(createBullet("Таблица расписания превращается в карточки: имя учителя — шапка, кабинет и класс — строкой ниже, уроки — карточками в две колонки с номером в углу."));
                children.push(createBullet("Замены выделяются цветной рамкой и значком 📌, пустые уроки — приглушённые и пунктирные."));
                children.push(createBullet("Журнал замен тоже превращается в карточки с крупными подписями «Отсутствует», «Предмет», «Заменяет»."));
                children.push(createBullet("Вкладки расписания показывают только иконки, растягиваются на всю ширину."));
                children.push(createBullet("Скрываются второстепенные элементы — переключатель темы, индикаторы, подсказки."));
                children.push(createBullet("Все поля ввода и кнопки становятся крупными (высота 40–42 px), удобными для нажатия пальцем."));
                children.push(createBullet("Панель фильтров и дней недели — вертикальная."));
                children.push(createBullet("Модальные окна — на всю ширину, кнопки — в столбик."));
                children.push(createBullet("Консоль администратора: таблица учителей — карточками, вкладки — горизонтальной прокруткой."));
                children.push(createBullet("Нет горизонтального скролла — всё влезает в ширину экрана телефона."));
                children.push(createBullet("Страница свободно прокручивается вниз — можно досмотреть весь контент."));
                children.push(createTextParagraph("Состояние мобильного режима сохраняется в браузере — при следующем открытии страницы он включится автоматически."));
                children.push(createTextParagraph("Чтобы вернуться к обычному виду — нажмите ту же кнопку (она сменится на «🖥 Обычная»)."));

                children.push(createHeading1("7. Пароль и безопасность"));
                children.push(createTextParagraph("Пароль по умолчанию — sever2. Он используется для:"));
                children.push(createBullet("входа в режим редактирования;"));
                children.push(createBullet("открытия консоли администратора;"));
                children.push(createBullet("доступа к защищённым вкладкам консоли (Безопасность, Конструктор, Обновление, Опасная зона)."));
                children.push(createTextParagraph("Пароль хранится в браузере (localStorage). Если очистить хранилище — вернётся стандартный пароль sever2."));
                children.push(createTextParagraph("Рекомендация: смените пароль сразу после первого входа через вкладку «🔐 Безопасность»."));

                children.push(createHeading1("8. Рекомендации по работе"));
                children.push(createBullet("Делайте экспорт JSON перед массовыми операциями и обновлением."));
                children.push(createBullet("Раз в неделю создавайте резервную копию через вкладку «💾 Импорт/Экспорт»."));
                children.push(createBullet("Перед обновлением HTML всегда проверяйте новый файл через встроенную валидацию."));
                children.push(createBullet("Не работайте в двух вкладках браузера одновременно — данные могут перезаписаться."));
                children.push(createBullet("После работы выходите из режима администратора кнопкой «🚪 Выйти»."));

                children.push(new Paragraph({ children: [], spacing: { before: 600 } }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: "Успешной работы с приложением!", bold: true, italics: true, size: 26, font: "Times New Roman", color: "1A3A6B" })],
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 200, after: 200 }
                }));

                const doc = new Document({
                    creator: "МОУ «Северная СОШ №2»",
                    title: "Полная инструкция · Расписание",
                    styles: { default: { document: { run: { font: "Times New Roman", size: 22, color: "1F2A44" } } } },
                    sections: [{
                        properties: { page: { margin: { top: 1134, right: 850, bottom: 1134, left: 1701 } } },
                        children: children
                    }]
                });

                const blob = await Packer.toBlob(doc);
                const url = URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url;
                a.download = `Полная_инструкция_${formatDateForFile()}.docx`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                if (statusEl) statusEl.textContent = "✅ Готово";
                showToast("📘 Полная инструкция скачана", "success");
                setTimeout(() => { if (statusEl) statusEl.textContent = ""; }, 3000);
            } catch (err) {
                console.error(err);
                if (statusEl) statusEl.textContent = `❌ Ошибка: ${err.message}`;
                showToast(`❌ Ошибка: ${err.message}`, "error");
            }
        }

        async function generateShortGuide() {
            const statusEl = document.getElementById("guideStatus");
            if (statusEl) statusEl.textContent = "⏳ Формирование...";
            try {
                const children = [];

                children.push(new Paragraph({ children: [], spacing: { before: 1000, after: 0 } }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: SCHOOL_NAME, bold: true, size: 44, font: "Times New Roman", color: "0B2A4A" })],
                    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 150 }
                }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: SCHOOL_SUBTITLE, size: 22, font: "Times New Roman", color: "3D5A7A" })],
                    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 800 }
                }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: "КРАТКАЯ ИНСТРУКЦИЯ", bold: true, size: 40, font: "Times New Roman", color: "1A3A6B" })],
                    alignment: AlignmentType.CENTER, spacing: { before: 400, after: 200 }
                }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: "Веб-приложение «Расписание учителей»", italics: true, size: 24, font: "Times New Roman", color: "3D5A7A" })],
                    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 400 }
                }));

                children.push(new Paragraph({ children: [new PageBreak()] }));

                children.push(createHeading1("🚀 Быстрый старт"));
                children.push(createNumberedItem("Откройте файл raspisanie.html в браузере (двойным щелчком).", 1));
                children.push(createNumberedItem("Наверху — четыре вкладки: 👨‍🏫 Учительское, 🎓 Детское, 📚 ИУП 5-8, 📌 Журнал замен. Переключайтесь между ними по клику.", 2));
                children.push(createNumberedItem("Чтобы редактировать, нажмите «✎ Редактировать» и введите пароль sever2.", 3));
                children.push(createNumberedItem("Для просмотра с телефона нажмите «📱 Мобильная» в шапке — расписание превратится в удобные карточки.", 4));
                children.push(createNumberedItem("После работы нажмите «🚪 Выйти», чтобы закрыть доступ.", 5));

                children.push(createHeading1("📚 Что делают вкладки"));
                children.push(createTextParagraph("👨‍🏫 Учительское — расписание по учителям. Здесь можно посмотреть, кто какой урок ведёт, и указать замену. При вводе замены приложение покажет список свободных учителей.", { bold: true }));
                children.push(createTextParagraph("🎓 Детское — расписание по классам: какой урок, какой учитель, в каком кабинете.", { bold: true }));
                children.push(createTextParagraph("📚 ИУП 5-8 — индивидуальные планы учеников с ОВЗ.", { bold: true }));
                children.push(createTextParagraph("📌 Журнал замен — отдельный журнал для бухгалтерии. Добавляйте записи о том, кто кого заменял.", { bold: true }));

                children.push(createHeading1("✎ Как редактировать расписание"));
                children.push(createNumberedItem("Нажмите «✎ Редактировать» и введите пароль sever2.", 1));
                children.push(createNumberedItem("Кликните по любой ячейке таблицы — откроется мини-редактор.", 2));
                children.push(createNumberedItem("Введите предмет в верхнее поле.", 3));
                children.push(createNumberedItem("Если нужна замена — впишите ФИО в поле «📌 Замена» (в режиме «Учительское» появится список свободных).", 4));
                children.push(createNumberedItem("Enter — сохранить, Esc — отменить.", 5));
                children.push(createNumberedItem("Нажмите «💾 Сохранить», чтобы сохранить изменения.", 6));
                children.push(createTextParagraph("Кнопка «↺ Сбросить» вернёт исходное расписание."));

                children.push(createHeading1("📌 Журнал замен"));
                children.push(createTextParagraph("Как добавить замену:", { bold: true }));
                children.push(createNumberedItem("Перейдите на вкладку «📌 Журнал замен».", 1));
                children.push(createNumberedItem("В форме «➕ Новая замена» укажите дату.", 2));
                children.push(createNumberedItem("Выберите тип урока: «Стандартный урок» или «Внеурочная деятельность».", 3));
                children.push(createNumberedItem("Выберите отсутствующего учителя и его предмет.", 4));
                children.push(createNumberedItem("Выберите заменяющего учителя и его предмет.", 5));
                children.push(createNumberedItem("Нажмите «➕ Добавить».", 6));
                children.push(createTextParagraph("Просмотр:", { bold: true }));
                children.push(createBullet("📅 Месяц — показать записи за месяц."));
                children.push(createBullet("📆 День — показать записи за конкретный день (кнопки ◀ ▶)."));
                children.push(createBullet("🗓 Календарь — включить мини-календарь, где подсвечены дни с заменами."));
                children.push(createTextParagraph("Экспорт:", { bold: true }));
                children.push(createBullet("📤 Excel — скачать журнал."));
                children.push(createBullet("🖨 Печать — распечатать с колонкой для подписи."));

                children.push(createHeading1("👤 Моё расписание"));
                children.push(createNumberedItem("Нажмите «👤 Моё расписание».", 1));
                children.push(createNumberedItem("Выберите учителя / класс / ученика из списка.", 2));
                children.push(createNumberedItem("Выберите день: «Сегодня», «Завтра», конкретный день или «Вся неделя».", 3));
                children.push(createNumberedItem("Нажмите «🖨 Распечатать» — распечатать.", 4));
                children.push(createNumberedItem("Кнопка «🔔 Звонки» — настроить время начала и конца уроков.", 5));

                children.push(createHeading1("📱 Мобильная версия"));
                children.push(createTextParagraph("Кнопка «📱 Мобильная» в шапке — для удобного просмотра с телефона."));
                children.push(createTextParagraph("Что меняется:", { bold: true }));
                children.push(createBullet("Расписание превращается в карточки: учитель — шапка, уроки — плитки в две колонки."));
                children.push(createBullet("Замены выделяются цветом и рамкой."));
                children.push(createBullet("Журнал замен — карточки с крупными подписями."));
                children.push(createBullet("Все кнопки и поля — крупные, удобные для нажатия пальцем."));
                children.push(createBullet("Нет горизонтального скролла — всё влезает в ширину экрана."));
                children.push(createBullet("Страница свободно прокручивается вниз."));
                children.push(createTextParagraph("Состояние сохраняется — при следующем открытии режим включится автоматически."));

                children.push(createHeading1("⚙ Консоль администратора"));
                children.push(createTextParagraph("Открывается кнопкой «⚙ Консоль», пароль sever2. Вкладки консоли:"));
                children.push(createBullet("👥 Учителя — добавление, удаление, редактирование."));
                children.push(createBullet("⚡ Массовые операции — очистка расписания."));
                children.push(createBullet("💾 Импорт/Экспорт — Excel, JSON, инструкции Word."));
                children.push(createBullet("📜 История — отмена изменений."));
                children.push(createBullet("🔐 Безопасность — смена пароля (нужен пароль)."));
                children.push(createBullet("🎨 Конструктор — смена дизайна (нужен пароль)."));
                children.push(createBullet("🔄 Обновление — загрузка нового HTML (нужен пароль)."));
                children.push(createBullet("⚠ Опасная зона — полный сброс (нужен пароль)."));

                children.push(createHeading1("🔄 Обновление приложения"));
                children.push(createNumberedItem("Откройте консоль → вкладка «🔄 Обновление сайта».", 1));
                children.push(createNumberedItem("Нажмите «💾 Бэкап данных» — создайте резервную копию.", 2));
                children.push(createNumberedItem("Перетащите новый HTML-файл в зону загрузки.", 3));
                children.push(createNumberedItem("Дождитесь проверки файла.", 4));
                children.push(createNumberedItem("Нажмите «✅ Применить обновление».", 5));
                children.push(createTextParagraph("Если что-то пошло не так — используйте «⏪ Откатить» или «📂 Восстановить из бэкапа»."));

                children.push(createHeading1("🛡️ Важные правила"));
                children.push(createBullet("Пароль по умолчанию — sever2. Смените его через «🔐 Безопасность»."));
                children.push(createBullet("Делайте резервные копии (JSON) перед массовыми операциями."));
                children.push(createBullet("Перед обновлением HTML всегда создавайте бэкап данных."));
                children.push(createBullet("Не работайте в двух вкладках браузера одновременно."));
                children.push(createBullet("Выходите из режима администратора кнопкой «🚪 Выйти»."));

                children.push(new Paragraph({ children: [], spacing: { before: 400 } }));
                children.push(new Paragraph({
                    children: [new TextRun({ text: "Успешной работы!", bold: true, italics: true, size: 24, font: "Times New Roman", color: "1A3A6B" })],
                    alignment: AlignmentType.CENTER
                }));

                const doc = new Document({
                    creator: "МОУ «Северная СОШ №2»",
                    title: "Краткая инструкция",
                    styles: { default: { document: { run: { font: "Times New Roman", size: 22, color: "1F2A44" } } } },
                    sections: [{
                        properties: { page: { margin: { top: 1134, right: 850, bottom: 1134, left: 1701 } } },
                        children: children
                    }]
                });

                const blob = await Packer.toBlob(doc);
                const url = URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url;
                a.download = `Краткая_инструкция_${formatDateForFile()}.docx`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                if (statusEl) statusEl.textContent = "✅ Готово";
                showToast("📄 Краткая инструкция скачана", "success");
                setTimeout(() => { if (statusEl) statusEl.textContent = ""; }, 3000);
            } catch (err) {
                console.error(err);
                if (statusEl) statusEl.textContent = `❌ Ошибка: ${err.message}`;
                showToast(`❌ Ошибка: ${err.message}`, "error");
            }
        }

        function initGuideButtons() {
            const btn1 = document.getElementById("downloadAdminGuideBtn");
            const btn2 = document.getElementById("downloadShortGuideBtn");
            if (btn1) btn1.addEventListener("click", generateAdminGuide);
            if (btn2) btn2.addEventListener("click", generateShortGuide);
        }

        function init() {
            initAlternativeSchedules();
            substitutionsJournal = loadSubstitutions();
            substitutionsHistory = loadSubstitutionsHistory();
            dictTeachersManual = loadDictTeachers();
            dictSubjectsManual = loadDictSubjects();
            initTheme();
            initMobileView();
            loadData();
            loadHistory();
            populateTeacherSelectForType();
            initFilters();
            initDayFilter();
            renderTable();
            initAnimatedBg();
            initConsoleTabs();
            initBulkActions();
            initImportExport();
            initExcelImport();
            initDangerZone();
            initHistoryPanel();
            initSubstitutionsTab();
            initConstructorTab();
            initUpdateTab();
            applyConstructorConfig();
            setTimeout(bindCustomTabClicks, 100);
            updateHistoryUI();
            updateSubstituteIndicator();

            document.getElementById("exportExcelBtn").addEventListener("click", exportToExcel);
            document.getElementById("exportExcelHistoryBtn").addEventListener("click", exportHistoryToExcel);

            document.getElementById("editToggleBtn").addEventListener("click", () => {
                if (editMode) exitEditMode();
                else openPasswordModal(
                    "🔒 Вход в режим редактирования",
                    "Введите пароль для изменения расписания.",
                    enterEditMode
                );
            });

            document.getElementById("consoleBtn").addEventListener("click", openConsole);
            document.getElementById("closeConsoleBtn").addEventListener("click", closeConsole);
            document.getElementById("consoleModal").addEventListener("click", (e) => {
                if (e.target.id === "consoleModal") closeConsole();
            });

            document.getElementById("confirmPasswordBtn").addEventListener("click", tryConfirmPassword);
            document.getElementById("cancelPasswordBtn").addEventListener("click", closePasswordModal);
            document.getElementById("passwordInput").addEventListener("keydown", (e) => {
                if (e.key === "Enter") { e.preventDefault(); tryConfirmPassword(); }
                if (e.key === "Escape") { e.preventDefault(); closePasswordModal(); }
            });
            document.getElementById("passwordModal").addEventListener("click", (e) => {
                if (e.target.id === "passwordModal") closePasswordModal();
            });

            document.getElementById("addTeacherBtn").addEventListener("click", () => openTeacherEditModal(-1));
            document.getElementById("cancelTeacherEditBtn").addEventListener("click", closeTeacherEditModal);
            document.getElementById("saveTeacherBtn").addEventListener("click", saveTeacherFromModal);
            document.getElementById("teacherEditModal").addEventListener("click", (e) => {
                if (e.target.id === "teacherEditModal") closeTeacherEditModal();
            });
            document.getElementById("teacherNameInput").addEventListener("keydown", (e) => {
                if (e.key === "Enter") { e.preventDefault(); saveTeacherFromModal(); }
                if (e.key === "Escape") { e.preventDefault(); closeTeacherEditModal(); }
            });

            document.getElementById("saveBtn").addEventListener("click", saveData);
            document.getElementById("resetBtn").addEventListener("click", resetData);

            document.getElementById("undoLastConsoleBtn").addEventListener("click", undoLastChange);
            document.getElementById("clearHistoryBtn").addEventListener("click", clearHistory);

            initGuideButtons();
            initMySchedule();
            initSecurityTab();

            document.getElementById("logoutAdminBtn").addEventListener("click", logoutAdmin);
            updateAdminModeUI();

            document.querySelectorAll(".schedule-tab").forEach(tab => {
                tab.addEventListener("click", () => {
                    if (tab.dataset.schedule && tab.dataset.schedule.startsWith("custom:")) return;
                    switchSchedule(tab.dataset.schedule);
                });
            });

            if (typeof currentScheduleType === "string" && currentScheduleType.startsWith("custom:")) {
                const tabId = currentScheduleType.split(":")[1];
                setTimeout(() => switchToCustomTab(tabId), 200);
            } else if (currentScheduleType !== "teachers") {
                document.body.classList.add("mode-" + currentScheduleType);
                document.querySelectorAll(".schedule-tab").forEach(t => {
                    t.classList.toggle("active", t.dataset.schedule === currentScheduleType);
                });
                const hintText = document.getElementById("scheduleHintText");
                if (hintText) {
                    if (currentScheduleType === "classes") {
                        hintText.innerHTML = "<strong>Расписание по классам</strong>: редактирование предметов и замен.";
                    } else if (currentScheduleType === "iup") {
                        hintText.innerHTML = "<strong>Индивидуальные учебные планы</strong>: редактирование занятий и замен.";
                    } else if (currentScheduleType === "substitutions") {
                        hintText.innerHTML = "<strong>Журнал замен учителей</strong>: дашборд с формой добавления и статистикой.";
                    }
                }
                const searchInput = document.getElementById("teacherSearch");
                if (searchInput) {
                    if (currentScheduleType === "classes") searchInput.placeholder = "Класс (например, 5а)...";
                    else if (currentScheduleType === "iup") searchInput.placeholder = "Имя ученика...";
                }
                populateTeacherSelectForType();
                if (currentScheduleType === "substitutions") {
                    renderSubstitutionsPanel();
                } else {
                    renderTable();
                }
            }
        }

        document.addEventListener("DOMContentLoaded", init);
    </script>
</body>
</html>
