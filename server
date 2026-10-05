const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const https = require('https');
const querystring = require('querystring');
const ExcelJS = require('exceljs');
const { Resend } = require('resend');
const { createClient } = require('@supabase/supabase-js');

const app = express();
const PORT = process.env.PORT || 3000;

// SUPABASE CONNECTION
const SUPABASE_URL = process.env.SUPABASE_URL || 'YOUR_SUPABASE_URL';
const SUPABASE_KEY = process.env.SUPABASE_KEY || 'YOUR_SUPABASE_SERVICE_ROLE_OR_ANON_KEY';
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

// AUTOMATICALLY GENERATE TEMPLATE.XLSX IF MISSING
async function ensureExcelTemplateExists() {
  const templatePath = path.join(__dirname, 'template.xlsx');
  
  if (fs.existsSync(templatePath)) {
    console.log('[EXCEL TEMPLATE] "template.xlsx" already exists.');
    return;
  }

  console.log('[EXCEL TEMPLATE] Generating professional attendance template...');
  
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Attendance Log');

  worksheet.columns = [
    { key: 'studentId', width: 18 },
    { key: 'name', width: 28 },
    { key: 'gradeSection', width: 20 },
    { key: 'position', width: 18 },
    { key: 'event', width: 24 },
    { key: 'scanType', width: 15 },
    { key: 'status', width: 15 },
    { key: 'duration', width: 15 },
    { key: 'timestamp', width: 25 }
  ];

  worksheet.mergeCells('B1:H1');
  worksheet.mergeCells('B2:H2');

  worksheet.getCell('B1').value = 'GENERAL ATTENDANCE MANAGEMENT SYSTEM';
  worksheet.getCell('B1').font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FF1B365D' } };
  worksheet.getCell('B1').alignment = { horizontal: 'center', vertical: 'middle' };

  worksheet.getCell('B2').value = 'OFFICIAL ATTENDANCE REPORT LOG';
  worksheet.getCell('B2').font = { name: 'Segoe UI', size: 11, italic: true, color: { argb: 'FF777777' } };
  worksheet.getCell('B2').alignment = { horizontal: 'center', vertical: 'middle' };

  const headers = [
    'ID NUMBER', 'FULL NAME', 'GROUP / SECTION', 
    'POSITION / ROLE', 'EVENT NAME', 'SCAN TYPE', 
    'STATUS', 'DURATION', 'TIMESTAMP'
  ];

  const headerRow = worksheet.getRow(4);
  headerRow.values = headers;
  headerRow.height = 26;

  headerRow.eachCell((cell) => {
    cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1B365D' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
      right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
    };
  });

  await workbook.xlsx.writeFile(templatePath);
  console.log('[EXCEL TEMPLATE] "template.xlsx" created successfully!');
}

async function getConfig() {
  const { data, error } = await supabase.from('config').select('*').eq('id', 1).single();
  if (error || !data) {
    const defaultConfig = {
      id: 1,
      system_name: 'General Attendance System',
      logo_path: '',
      events: ['General Event', 'Orientation', 'Meeting', 'Seminar'],
      current_event: 'General Event',
      cutoff_time: '08:00',
      latest_uid: '',
      enable_email: false,
      gmail_user: '',
      gmail_pass: '',
      enable_sms: false,
      semaphore_api_key: ''
    };
    await supabase.from('config').upsert([defaultConfig]);
    return defaultConfig;
  }
  return data;
}

// UPLOADS SETUP FOR LOGOS & PARTICIPANT PHOTOS
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`)
});

const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith('image/')) cb(null, true);
    else cb(new Error('Only image files are allowed!'), false);
  }
});

// MIDDLEWARES
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use('/uploads', express.static(uploadsDir));

// NOTIFICATIONS
const resend = new Resend(process.env.RESEND_API_KEY || 'YOUR_RESEND_API_KEY');

async function sendEmailNotification(recipientEmail, studentName, scanType, status, eventName, timestamp, duration) {
  if (!recipientEmail) return;
  const durationText = duration ? `<li><strong>Duration:</strong> ${duration}</li>` : '';
  try {
    await resend.emails.send({
      from: 'Attendance System <onboarding@resend.dev>',
      to: recipientEmail,
      subject: `[${scanType}] Attendance Alert: ${studentName}`,
      html: `
        <div style="font-family: Arial, sans-serif; padding: 15px; border: 1px solid #ddd; border-radius: 6px;">
          <h2 style="color: #2c3e50;">Attendance Notification (${scanType})</h2>
          <p>Hello,</p>
          <p>This is to confirm that <strong>${studentName}</strong> logged <strong>${scanType}</strong>.</p>
          <ul>
            <li><strong>Event:</strong> ${eventName}</li>
            <li><strong>Scan Type:</strong> <span style="color:#2980b9; font-weight:bold;">${scanType}</span></li>
            <li><strong>Status:</strong> <span style="color:${status === 'LATE' ? '#e74c3c' : '#2ecc71'}; font-weight:bold;">${status}</span></li>
            <li><strong>Time:</strong> ${timestamp}</li>
            ${durationText}
          </ul>
        </div>
      `
    });
  } catch (error) {
    console.error('[EMAIL ERROR]', error.message);
  }
}

function calculateDuration(timeInDate, timeOutDate) {
  const diffMs = timeOutDate - timeInDate;
  if (isNaN(diffMs) || diffMs < 0) return 'N/A';
  const totalMinutes = Math.floor(diffMs / (1000 * 60));
  const hours = Math.floor(totalMinutes / 60);
  const mins = totalMinutes % 60;
  return hours > 0 ? `${hours}h ${mins}m` : `${mins}m`;
}

// API: ESP8266 / HARDWARE SCANNER ENDPOINT
app.post('/api/scan', async (req, res) => {
  try {
    const { uid } = req.body;
    if (!uid) return res.status(400).json({ status: 'error', message: 'No UID provided' });

    const cleanUid = uid.trim().toUpperCase();
    const config = await getConfig();
    
    await supabase.from('config').update({ latest_uid: cleanUid }).eq('id', 1);

    const { data: student } = await supabase
      .from('students')
      .select('*')
      .eq('uid', cleanUid)
      .single();

    const now = new Date();

    if (student) {
      const eventName = student.assigned_event || config.current_event || 'General Event';
      const startOfDay = new Date(now);
      startOfDay.setHours(0, 0, 0, 0);

      const { data: lastLogs } = await supabase
        .from('attendance')
        .select('*')
        .eq('uid', cleanUid)
        .eq('event', eventName)
        .gte('raw_timestamp', startOfDay.toISOString())
        .order('raw_timestamp', { ascending: false })
        .limit(1);

      const lastLog = lastLogs && lastLogs.length > 0 ? lastLogs[0] : null;

      let scanType = 'TIME-IN';
      let duration = '';
      let statusLabel = 'ON TIME';

      if (lastLog && lastLog.scan_type === 'TIME-IN') {
        scanType = 'TIME-OUT';
        statusLabel = 'COMPLETED';
        duration = calculateDuration(new Date(lastLog.raw_timestamp), now);
      } else {
        const currentTimeStr = now.toTimeString().slice(0, 5);
        statusLabel = currentTimeStr > (config.cutoff_time || '08:00') ? 'LATE' : 'ON TIME';
      }

      const record = {
        uid: cleanUid,
        name: student.name,
        student_id: student.student_id,
        year_level: student.year_level || 'N/A',
        section: student.section || 'N/A',
        position: student.position || 'Member',
        email: student.email || '',
        phone: student.phone || '',
        event: eventName,
        scan_type: scanType,
        status: statusLabel,
        duration: duration || 'N/A',
        timestamp: now.toLocaleString(),
        raw_timestamp: now.toISOString()
      };

      await supabase.from('attendance').insert([record]);

      if (config.enable_email && student.email) {
        sendEmailNotification(student.email, student.name, scanType, statusLabel, eventName, record.timestamp, duration);
      }

      return res.json({ 
        status: 'success', 
        scanType, 
        isLate: statusLabel === 'LATE', 
        message: `${scanType} recorded for ${student.name}`,
        student: {
          name: student.name,
          studentId: student.student_id,
          yearLevel: student.year_level,
          section: student.section,
          position: student.position,
          photoUrl: student.photo_url || ''
        }
      });
    } else {
      return res.json({ status: 'unknown', message: 'Card not registered', uid: cleanUid });
    }
  } catch (err) {
    console.error('[SCAN ERROR]', err);
    res.status(500).json({ status: 'error', message: 'Server Error' });
  }
});

// REALTIME DATA ENDPOINT FOR DASHBOARD UI
app.get('/api/live-data', async (req, res) => {
  try {
    const config = await getConfig();
    const { data: students } = await supabase.from('students').select('*').order('name', { ascending: true });
    const { data: attendance } = await supabase.from('attendance').select('*').order('raw_timestamp', { ascending: false }).limit(100);

    // Also attach photo URLs if matching by UID for live scanned display
    res.json({
      latestUid: config.latest_uid || '',
      attendance: attendance || [],
      students: students || []
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// EXPORT TO EXCEL
app.get('/api/export-excel', async (req, res) => {
  try {
    await ensureExcelTemplateExists();

    const selectedEvent = req.query.event;
    let query = supabase.from('attendance').select('*').order('raw_timestamp', { ascending: false });
    if (selectedEvent && selectedEvent !== 'ALL') {
      query = query.eq('event', selectedEvent);
    }

    const { data: attendance } = await query;

    const templatePath = path.join(__dirname, 'template.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const worksheet = workbook.getWorksheet(1);

    let startRow = 5;

    (attendance || []).forEach((row, index) => {
      const currentRow = worksheet.getRow(startRow + index);
      
      currentRow.getCell(1).value = row.student_id;
      currentRow.getCell(2).value = row.name;
      currentRow.getCell(3).value = `${row.year_level || ''} - ${row.section || ''}`;
      currentRow.getCell(4).value = row.position || 'Member';
      currentRow.getCell(5).value = row.event;
      currentRow.getCell(6).value = row.scan_type || 'TIME-IN';
      currentRow.getCell(7).value = row.status;
      currentRow.getCell(8).value = row.duration || 'N/A';
      currentRow.getCell(9).value = row.timestamp;

      currentRow.commit();
    });

    const filename = selectedEvent && selectedEvent !== 'ALL' 
      ? `Attendance_${selectedEvent.replace(/\s+/g, '_')}.xlsx` 
      : 'Attendance_Report.xlsx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error('[EXCEL EXPORT ERROR]', err.message);
    res.status(500).send('Error generating Excel file: ' + err.message);
  }
});

// SETTINGS & ADMIN ENDPOINTS
app.post('/api/update-system-name', async (req, res) => {
  const { systemName } = req.body;
  if (systemName) {
    await supabase.from('config').update({ system_name: systemName }).eq('id', 1);
  }
  res.redirect('/');
});

app.post('/api/upload-logo', upload.single('logoFile'), async (req, res) => {
  if (req.file) {
    const logoPath = `/uploads/${req.file.filename}`;
    await supabase.from('config').update({ logo_path: logoPath }).eq('id', 1);
  }
  res.redirect('/');
});

app.post('/api/remove-logo', async (req, res) => {
  await supabase.from('config').update({ logo_path: '' }).eq('id', 1);
  res.redirect('/');
});

app.post('/api/notification-settings', async (req, res) => {
  const { enableEmail, gmailUser, gmailPass, enableSms, semaphoreApiKey } = req.body;
  await supabase.from('config').update({
    enable_email: enableEmail === 'on',
    gmail_user: gmailUser || '',
    gmail_pass: gmailPass && gmailPass !== '******' ? gmailPass : undefined,
    enable_sms: enableSms === 'on',
    semaphore_api_key: semaphoreApiKey || ''
  }).eq('id', 1);
  res.redirect('/');
});

app.post('/api/event-settings', async (req, res) => {
  const { newEvent, activeEvent, cutoffTime } = req.body;
  const config = await getConfig();
  let events = Array.isArray(config.events) ? config.events : ['General Event'];

  if (newEvent && !events.includes(newEvent)) {
    events.push(newEvent);
    config.current_event = newEvent;
  } else if (activeEvent) {
    config.current_event = activeEvent;
  }

  const updatePayload = { events, current_event: config.current_event };
  if (cutoffTime) updatePayload.cutoff_time = cutoffTime;

  await supabase.from('config').update(updatePayload).eq('id', 1);
  res.redirect('/');
});

app.post('/api/delete-event', async (req, res) => {
  const { eventToDelete } = req.body;
  if (eventToDelete) {
    const config = await getConfig();
    let events = (config.events || []).filter(e => e !== eventToDelete);
    let currentEvent = config.current_event;
    if (currentEvent === eventToDelete) {
      currentEvent = events[0] || 'General Event';
    }
    await supabase.from('config').update({ events, current_event: currentEvent }).eq('id', 1);
  }
  res.redirect('/');
});

// REGISTER / EDIT PARTICIPANT WITH PHOTO SUPPORT
app.post('/api/register', upload.single('photoFile'), async (req, res) => {
  try {
    const { studentIdId, uid, name, studentId, yearLevel, section, assignedEvent, position, customPosition, email, phone } = req.body;
    let finalPosition = (position === 'Other' && customPosition) ? customPosition.trim() : position || 'Member';
    const cleanUid = uid ? uid.trim().toUpperCase() : '';

    let photoUrl = req.body.existingPhotoUrl || '';
    if (req.file) {
      photoUrl = `/uploads/${req.file.filename}`;
    }

    const payload = {
      uid: cleanUid,
      name,
      student_id: studentId,
      year_level: yearLevel || 'Grade 7',
      section: section || 'A',
      position: finalPosition,
      email: email || '',
      phone: phone || '',
      assigned_event: assignedEvent || 'General Event',
      photo_url: photoUrl
    };

    if (studentIdId) {
      await supabase.from('students').update(payload).eq('id', studentIdId);
    } else {
      await supabase.from('students').insert([payload]);
    }
  } catch (err) {
    console.error('[SAVE ERROR]', err.message);
  }
  res.redirect('/');
});

app.post('/api/delete-student', async (req, res) => {
  const { id } = req.body;
  if (id) await supabase.from('students').delete().eq('id', id);
  res.redirect('/');
});

app.post('/api/clear-logs', async (req, res) => {
  await supabase.from('attendance').delete().neq('id', '00000000-0000-0000-0000-000000000000');
  res.redirect('/');
});

// PUBLIC REGISTRATION FORM
app.get('/register', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Participant Registration</title>
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #f4f6f9; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 20px 0; }
        .card { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); width: 100%; max-width: 420px; box-sizing: border-box; }
        h2 { text-align: center; color: #2c3e50; margin-bottom: 20px; }
        label { font-weight: bold; font-size: 14px; color: #34495e; }
        input, select { width: 100%; padding: 10px; margin: 6px 0 16px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
        button { width: 100%; background: #27ae60; color: white; padding: 12px; border: none; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 16px; margin-top: 10px; }
        button:hover { background: #219150; }
        .hidden { display: none; }
      </style>
    </head>
    <body>
      <div class="card">
        <h2>Participant Registration</h2>
        <form action="/api/public-register" method="POST" enctype="multipart/form-data">
          <label>ID Number / Badge ID:</label>
          <input type="text" name="studentId" placeholder="e.g. 2026-0001" required>

          <label>Full Name:</label>
          <input type="text" name="name" placeholder="John Doe" required>

          <label>Group / Section / Department:</label>
          <input type="text" name="section" placeholder="e.g. IT Department / Section A" required>

          <label>Position / Role:</label>
          <select name="position" id="posSelect" onchange="checkCustom()" required>
            <option value="Member">Member</option>
            <option value="Officer">Officer</option>
            <option value="Staff">Staff</option>
            <option value="Guest">Guest</option>
            <option value="Other">Custom Position...</option>
          </select>

          <div id="customBox" class="hidden">
            <label>Specify Custom Role:</label>
            <input type="text" name="customPosition" placeholder="Enter role">
          </div>

          <label>Email Address:</label>
          <input type="email" name="email" placeholder="john@example.com" required>

          <label>Profile Picture:</label>
          <input type="file" name="photoFile" accept="image/*">

          <button type="submit">Submit Registration</button>
        </form>
      </div>
      <script>
        function checkCustom() {
          const val = document.getElementById('posSelect').value;
          document.getElementById('customBox').classList.toggle('hidden', val !== 'Other');
        }
      </script>
    </body>
    </html>
  `);
});

app.post('/api/public-register', upload.single('photoFile'), async (req, res) => {
  try {
    const { name, email, studentId, section, position, customPosition } = req.body;
    let finalPosition = (position === 'Other' && customPosition) ? customPosition.trim() : position || 'Member';
    let photoUrl = req.file ? `/uploads/${req.file.filename}` : '';

    await supabase.from('students').insert([{
      name,
      email,
      student_id: studentId,
      section: section || 'General',
      position: finalPosition,
      photo_url: photoUrl,
      uid: ''
    }]);

    res.send(`
      <div style="text-align:center; padding:50px; font-family:Arial;">
        <h2 style="color:#2ecc71;">Registration Successful!</h2>
        <p>Thank you <strong>${name}</strong>! Your information has been registered.</p>
        <a href="/register" style="color:#2980b9; font-weight:bold; text-decoration:none;">Register another participant</a>
      </div>
    `);
  } catch (err) {
    res.status(500).send('Error: ' + err.message);
  }
});

// REDESIGNED GENERAL ADMIN DASHBOARD ( / )
app.get('/', async (req, res) => {
  const config = await getConfig();
  const eventList = Array.isArray(config.events) ? config.events : ['General Event'];
  const eventOptions = eventList.map(e => `<option value="${e}" ${e === config.current_event ? 'selected' : ''}>${e}</option>`).join('');

  const logoHtml = config.logo_path ? `<img src="${config.logo_path}" alt="System Logo" class="header-logo">` : '';

  res.send(`
  <!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${config.system_name || 'General Attendance System'}</title>
    <style>
      :root {
        --primary: #2563eb;
        --primary-dark: #1d4ed8;
        --success: #10b981;
        --danger: #ef4444;
        --warning: #f59e0b;
        --bg: #f8fafc;
        --card-bg: #ffffff;
        --text: #1e293b;
        --border: #e2e8f0;
      }
      body { font-family: 'Segoe UI', system-ui, sans-serif; margin: 0; padding: 20px; background: var(--bg); color: var(--text); }
      .header-container { display: flex; align-items: center; justify-content: space-between; background: var(--card-bg); padding: 15px 25px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 20px; }
      .header-left { display: flex; align-items: center; gap: 15px; }
      .header-logo { height: 50px; width: auto; object-fit: contain; border-radius: 6px; }
      h1, h2, h3 { color: var(--text); margin: 0; }
      .grid-layout { display: grid; grid-template-columns: 1fr 380px; gap: 20px; }
      @media(max-width: 1024px) { .grid-layout { grid-template-columns: 1fr; } }
      .card { background: var(--card-bg); padding: 20px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); margin-bottom: 20px; }
      input, select { width: 100%; padding: 10px; margin: 6px 0 14px 0; border: 1px solid var(--border); border-radius: 6px; box-sizing: border-box; }
      button, input[type="submit"] { background: var(--primary); color: white; padding: 10px 16px; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; }
      button:hover, input[type="submit"]:hover { background: var(--primary-dark); }
      .btn-danger { background: var(--danger); }
      .btn-danger:hover { background: #dc2626; }
      .btn-warning { background: var(--warning); }
      .btn-warning:hover { background: #d97706; }
      table { width: 100%; border-collapse: collapse; margin-top: 10px; }
      th, td { border-bottom: 1px solid var(--border); padding: 12px; text-align: left; font-size: 14px; }
      th { background: #f1f5f9; font-weight: 600; color: #475569; }
      .badge { padding: 4px 10px; border-radius: 9999px; font-weight: 600; font-size: 12px; display: inline-block; }
      .badge-ontime { background: #d1fae5; color: #065f46; }
      .badge-late { background: #fee2e2; color: #991b1b; }
      .badge-in { background: #e0f2fe; color: #0369a1; }
      .badge-out { background: #f3e8ff; color: #6b21a8; }
      .scanner-display { background: #0f172a; color: white; padding: 25px; border-radius: 12px; text-align: center; margin-bottom: 20px; display: flex; align-items: center; gap: 25px; }
      .scanner-avatar { width: 120px; height: 120px; border-radius: 50%; object-fit: cover; border: 4px solid var(--primary); background: #334155; }
      .scanner-info { text-align: left; flex: 1; }
      .nav-links { margin-bottom: 15px; font-size: 14px; }
      .nav-links a { color: var(--primary); text-decoration: none; font-weight: 600; }
    </style>
  </head>
  <body>

    <div class="header-container">
      <div class="header-left">
        ${logoHtml}
        <h1>${config.system_name || 'General Attendance System'}</h1>
      </div>
      <div class="nav-links">
        <a href="/register" target="_blank">📋 Public Registration Form</a>
      </div>
    </div>

    <!-- LIVE RFID SCANNER VISUAL & AUDIO FEEDBACK -->
    <div class="scanner-display" id="scannerCard">
      <img id="livePhoto" src="/uploads/default-avatar.png" class="scanner-avatar" onerror="this.src='data:image/svg+xml;utf8,<svg xmlns=\'http://www.w3.org/2000/svg\' width=\'120\' height=\'120\'><rect width=\'100%\' height=\'100%\' fill=\'%23334155\'/><text x=\'50%\' y=\'50%\' fill=\'%2394a3b8\' dominant-baseline=\'middle\' text-anchor=\'middle\' font-size=\'14\'>No Photo</text></svg>'">
      <div class="scanner-info">
        <div style="font-size: 13px; text-transform: uppercase; letter-spacing: 1px; color: #94a3b8; margin-bottom: 4px;">Latest RFID Scan Feedback</div>
        <h2 id="liveName" style="color: #ffffff; font-size: 26px; margin-bottom: 5px;">Waiting for RFID scan...</h2>
        <p id="liveDetails" style="color: #cbd5e1; margin: 0; font-size: 15px;">Scan any card to instantly display info and announce name.</p>
        <div style="margin-top: 8px;"><span id="liveBadge" class="badge badge-in">ID: <span id="liveUidDisplay">${config.latest_uid || 'None'}</span></span></div>
      </div>
    </div>

    <div class="grid-layout">
      <!-- LEFT COLUMN: TABLES -->
      <div>
        <div class="card">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
            <h2>Live Attendance Logs</h2>
            <div style="display: flex; gap: 10px; align-items: center;">
              <select id="exportEventSelect" style="width: auto; margin:0;">
                <option value="ALL">All Events</option>
                ${eventOptions}
              </select>
              <button onclick="downloadExcel()">Export Excel</button>
              <form action="/api/clear-logs" method="POST" onsubmit="return confirm('Clear all attendance logs?');" style="margin:0;">
                <button type="submit" class="btn-danger">Clear Logs</button>
              </form>
            </div>
          </div>
          <div style="overflow-x: auto;">
            <table>
              <thead>
                <tr>
                  <th>Participant</th>
                  <th>ID Number</th>
                  <th>Department / Section</th>
                  <th>Event</th>
                  <th>Type</th>
                  <th>Status</th>
                  <th>Duration</th>
                  <th>Time</th>
                </tr>
              </thead>
              <tbody id="attendanceTableBody"></tbody>
            </table>
          </div>
        </div>

        <div class="card">
          <h2>Registered Participants Database</h2>
          <div style="overflow-x: auto;">
            <table>
              <thead>
                <tr>
                  <th>Photo</th>
                  <th>ID Number</th>
                  <th>Name</th>
                  <th>Section</th>
                  <th>Position</th>
                  <th>Assigned Event</th>
                  <th>Card UID</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody id="studentsTableBody"></tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- RIGHT COLUMN: SETTINGS & REGISTRATION -->
      <div>
        <div class="card">
          <h3 id="formTitle">Register / Edit Participant</h3>
          <form action="/api/register" method="POST" enctype="multipart/form-data" id="registerForm" style="margin-top: 15px;">
            <input type="hidden" id="studentIdIdInput" name="studentIdId">
            <input type="hidden" id="existingPhotoUrlInput" name="existingPhotoUrl">

            <label><strong>RFID Card UID:</strong></label>
            <div style="display: flex; gap: 8px;">
              <input type="text" id="uidInput" name="uid" placeholder="Scan or type UID">
              <button type="button" onclick="useLatestUid()" style="margin-top: 6px; white-space: nowrap;">Get Last UID</button>
            </div>

            <label><strong>ID Number:</strong></label>
            <input type="text" id="studentIdInput" name="studentId" placeholder="e.g. 2026-001" required>

            <label><strong>Full Name:</strong></label>
            <input type="text" id="nameInput" name="name" placeholder="Full Name" required>

            <label><strong>Group / Section:</strong></label>
            <input type="text" id="sectionInput" name="section" placeholder="e.g. Section A / IT Dept" required>

            <label><strong>Position / Role:</strong></label>
            <select id="positionSelect" name="position">
              <option value="Member">Member</option>
              <option value="Officer">Officer</option>
              <option value="Staff">Staff</option>
              <option value="Guest">Guest</option>
            </select>

            <label><strong>Assign Event:</strong></label>
            <select id="eventSelect" name="assignedEvent">${eventOptions}</select>

            <label><strong>Profile Picture:</strong></label>
            <input type="file" name="photoFile" accept="image/*">

            <div style="display: flex; gap: 10px; margin-top: 10px;">
              <input type="submit" id="submitBtn" value="Save Participant" style="flex: 1; margin: 0;">
              <button type="button" id="cancelEditBtn" onclick="resetForm()" class="btn-danger" style="display: none; flex: 1;">Cancel</button>
            </div>
          </form>
        </div>

        <div class="card">
          <h3>System & Event Settings</h3>
          <form action="/api/update-system-name" method="POST" style="margin-top: 10px;">
            <label>System Title:</label>
            <input type="text" name="systemName" value="${config.system_name || 'General Attendance System'}">
            <input type="submit" value="Update Title">
          </form>

          <hr style="border: 0; border-top: 1px solid var(--border); margin: 15px 0;">

          <form action="/api/event-settings" method="POST">
            <label>Active Event:</label>
            <select name="activeEvent">${eventOptions}</select>

            <label>Add New Event:</label>
            <input type="text" name="newEvent" placeholder="Event Name">

            <label>Late Cut-off Time:</label>
            <input type="time" name="cutoffTime" value="${config.cutoff_time || '08:00'}">

            <input type="submit" value="Save Event Settings">
          </form>
        </div>
      </div>
    </div>

    <script>
      let registeredStudents = [];
      let lastSpokenUid = '';

      function speakName(text) {
        if ('speechSynthesis' in window) {
          window.speechSynthesis.cancel(); // Stop any pending speech
          const utterance = new SpeechSynthesisUtterance(text);
          utterance.rate = 1.0;
          utterance.pitch = 1.0;
          window.speechSynthesis.speak(utterance);
        }
      }

      function editStudent(id) {
        const st = registeredStudents.find(s => s.id === id);
        if (!st) return;

        document.getElementById('studentIdIdInput').value = st.id;
        document.getElementById('uidInput').value = st.uid || '';
        document.getElementById('studentIdInput').value = st.student_id;
        document.getElementById('nameInput').value = st.name;
        document.getElementById('sectionInput').value = st.section || '';
        if (st.position) document.getElementById('positionSelect').value = st.position;
        if (st.assigned_event) document.getElementById('eventSelect').value = st.assigned_event;
        document.getElementById('existingPhotoUrlInput').value = st.photo_url || '';

        document.getElementById('formTitle').innerText = 'Edit Participant (' + st.name + ')';
        document.getElementById('submitBtn').value = 'Update Participant';
        document.getElementById('cancelEditBtn').style.display = 'block';
        window.scrollTo({ top: 0, behavior: 'smooth' });
      }

      function resetForm() {
        document.getElementById('registerForm').reset();
        document.getElementById('studentIdIdInput').value = '';
        document.getElementById('existingPhotoUrlInput').value = '';
        document.getElementById('formTitle').innerText = 'Register / Edit Participant';
        document.getElementById('submitBtn').value = 'Save Participant';
        document.getElementById('cancelEditBtn').style.display = 'none';
      }

      async function updateDashboard() {
        try {
          const res = await fetch('/api/live-data');
          const data = await res.json();
          registeredStudents = data.students || [];

          if (data.latestUid) {
            document.getElementById('liveUidDisplay').innerText = data.latestUid;

            // Find latest attendance or student corresponding to this UID
            const latestScan = data.attendance && data.attendance.length > 0 ? data.attendance[0] : null;
            if (latestScan && latestScan.uid === data.latestUid) {
              const matchedStudent = registeredStudents.find(s => s.uid === data.latestUid);
              const photo = matchedStudent && matchedStudent.photo_url ? matchedStudent.photo_url : '/uploads/default-avatar.png';
              
              document.getElementById('livePhoto').src = photo;
              document.getElementById('liveName').innerText = latestScan.name + ' (' + latestScan.scan_type + ')';
              document.getElementById('liveDetails').innerText = 'ID: ' + latestScan.student_id + ' | ' + latestScan.section + ' | Status: ' + latestScan.status + ' | ' + latestScan.timestamp;

              // Speak name aloud if new UID scan detected
              if (lastSpokenUid !== data.latestUid + '-' + latestScan.timestamp) {
                lastSpokenUid = data.latestUid + '-' + latestScan.timestamp;
                speakName(latestScan.name + " " + latestScan.scan_type);
              }
            }
          }

          // Render Attendance Table
          const tbody = document.getElementById('attendanceTableBody');
          if (!data.attendance || data.attendance.length === 0) {
            tbody.innerHTML = '<tr><td colspan="8" style="text-align:center; color:#94a3b8;">No attendance records found.</td></tr>';
          } else {
            tbody.innerHTML = data.attendance.map(row => {
              const typeBadge = row.scan_type === 'TIME-OUT' ? '<span class="badge badge-out">TIME-OUT</span>' : '<span class="badge badge-in">TIME-IN</span>';
              const statusBadge = row.status === 'LATE' ? '<span class="badge badge-late">LATE</span>' : '<span class="badge badge-ontime">' + row.status + '</span>';
              return \`
                <tr>
                  <td><strong>\${row.name}</strong></td>
                  <td>\${row.student_id}</td>
                  <td>\${row.section}</td>
                  <td>\${row.event}</td>
                  <td>\${typeBadge}</td>
                  <td>\${statusBadge}</td>
                  <td><strong>\${row.duration || 'N/A'}</strong></td>
                  <td>\${row.timestamp}</td>
                </tr>
              \`;
            }).join('');
          }

          // Render Students Table
          const stBody = document.getElementById('studentsTableBody');
          if (!registeredStudents || registeredStudents.length === 0) {
            stBody.innerHTML = '<tr><td colspan="8" style="text-align:center; color:#94a3b8;">No participants registered.</td></tr>';
          } else {
            stBody.innerHTML = registeredStudents.map(st => {
              const imgTag = st.photo_url ? \`<img src="\${st.photo_url}" style="width:40px;height:40px;border-radius:50%;object-fit:cover;">\` : '<span style="color:#94a3b8;font-size:12px;">No Image</span>';
              return \`
                <tr>
                  <td>\${imgTag}</td>
                  <td>\${st.student_id}</td>
                  <td><strong>\${st.name}</strong></td>
                  <td>\${st.section || 'N/A'}</td>
                  <td><span style="color:var(--primary); font-weight:600;">\${st.position || 'Member'}</span></td>
                  <td>\${st.assigned_event || 'General Event'}</td>
                  <td>\${st.uid ? '<code>' + st.uid + '</code>' : '<span style="color:#f59e0b;font-weight:600;">Unlinked</span>'}</td>
                  <td>
                    <button type="button" class="btn-warning" onclick="editStudent('\${st.id}')" style="padding:6px 10px; font-size:12px;">Edit</button>
                    <form action="/api/delete-student" method="POST" style="display:inline;" onsubmit="return confirm('Delete participant?');">
                      <input type="hidden" name="id" value="\$.id}">
                      <button type="submit" class="btn-danger" style="padding:6px 10px; font-size:12px;">Delete</button>
                    </form>
                  </td>
                </tr>
              \`;
            }).join('');
          }

        } catch (err) {}
      }

      function useLatestUid() {
        const uid = document.getElementById('liveUidDisplay').innerText;
        if (uid && uid !== 'None') document.getElementById('uidInput').value = uid;
      }

      function downloadExcel() {
        const selected = document.getElementById('exportEventSelect').value;
        window.location.href = '/api/export-excel?event=' + encodeURIComponent(selected);
      }

      updateDashboard();
      setInterval(updateDashboard, 2000);
    </script>
  </body>
  </html>
  `);
});

// START SERVER
app.listen(PORT, async () => {
  await ensureExcelTemplateExists();
  console.log(`General Attendance Server running on port ${PORT}`);
});
