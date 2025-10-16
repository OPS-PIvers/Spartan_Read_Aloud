
let adminSessionToken = null;
let adminName = null;
let userRole = 'student'; // Default, will be updated on successful login
let allAssessments = [];

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function parseEmailList(input) {
  if (!input || typeof input !== 'string') {
    return '';
  }

  // Email regex pattern (basic but robust)
  const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;

  // Extract all email addresses from the input
  const emailMatches = input.match(emailRegex);

  if (!emailMatches || emailMatches.length === 0) {
    return '';
  }

  // Normalize: lowercase, trim, deduplicate
  const uniqueEmails = [...new Set(
    emailMatches.map(email => email.toLowerCase().trim())
  )];

  // Return as comma-separated string
  return uniqueEmails.join(', ');
}

(function checkStoredAdminSession() {
  const storedToken = localStorage.getItem('adminToken');
  const storedName = localStorage.getItem('adminName');
  const storedRole = localStorage.getItem('userRole');

  if (storedToken && storedName) {
    console.log('Found stored staff session, restoring...');
    adminSessionToken = storedToken;
    adminName = storedName;
    userRole = storedRole || 'teacher'; // Default to least-privileged role for safety
    // Show admin dashboard automatically
    showAdminDashboard();
  }
})();

function showAdminDashboard() {
  // Show admin dashboard
  const adminDashboard = document.getElementById('admin-dashboard-container');
  adminDashboard.style.display = 'block';

  // Set admin name
  document.getElementById('admin-name-display').textContent = `Welcome, ${adminName}`;

  // Load all assessments
  loadAllAssessments();
}

function logoutAdmin() {
  // Clear admin session
  adminSessionToken = null;
  adminName = null;
  allAssessments = [];

  // Clear stored session from localStorage
  localStorage.removeItem('adminToken');
  localStorage.removeItem('adminName');
  localStorage.removeItem('userRole');

  google.script.run.withSuccessHandler(function(html) {
    document.getElementById('app-container').innerHTML = html;
  }).getLoginView();
}

function showUploadTab(tab) {
  const fileTab = document.getElementById('upload-tab-file');
  const gdocTab = document.getElementById('upload-tab-gdoc');
  const fileSection = document.getElementById('upload-section-file');
  const gdocSection = document.getElementById('upload-section-gdoc');

  if (tab === 'file') {
    fileTab.style.borderBottomColor = '#1a73e8';
    fileTab.style.color = '#1a73e8';
    fileTab.style.fontWeight = '600';
    gdocTab.style.borderBottomColor = 'transparent';
    gdocTab.style.color = '#5f6368';
    gdocTab.style.fontWeight = '500';
    fileSection.style.display = 'block';
    gdocSection.style.display = 'none';
  } else {
    gdocTab.style.borderBottomColor = '#1a73e8';
    gdocTab.style.color = '#1a73e8';
    gdocTab.style.fontWeight = '600';
    fileTab.style.borderBottomColor = 'transparent';
    fileTab.style.color = '#5f6368';
    fileTab.style.fontWeight = '500';
    gdocSection.style.display = 'block';
    fileSection.style.display = 'none';
  }
}

function loadAllAssessments() {
  console.log('Loading all assessments for admin...');
  google.script.run
    .withSuccessHandler(onAllAssessmentsLoaded)
    .withFailureHandler((error) => {
      console.error('Error loading assessments:', error);
      alert('Failed to load assessments: ' + error.message);
      // If there's an error, clear stored session and show login
      logoutAdmin();
    })
    .getAllAssessments(adminSessionToken);
}

function onAllAssessmentsLoaded(result) {
  console.log('Staff assessments loaded:', result);

  if (result.error) {
    console.error('Error loading assessments:', result.error);

    // Check if error is due to unauthorized/expired token
    if (result.error.includes('Unauthorized') || result.error.includes('Staff access required') || result.error.includes('Admin access required')) {
      alert('Your session has expired. Please log in again.');
      logoutAdmin();
    } else {
      alert('Error: ' + result.error);
    }
    return;
  }

  // Update userRole if backend returned it (more reliable than localStorage)
  if (result.userRole) {
    userRole = result.userRole;
    localStorage.setItem('userRole', userRole);
    console.log(`User role: ${userRole}`);
  }

  console.log(`Successfully loaded ${result.assessments.length} assessments for ${userRole}`);
  allAssessments = result.assessments;
  renderAssessmentTable();
}

function renderAssessmentTable() {
  const tbody = document.getElementById('admin-assessment-tbody');
  const noAssessmentsMsg = document.getElementById('no-assessments-message');
  const addAssessmentSection = document.getElementById('add-assessment-section');

  // Show/hide assessment upload section based on role
  if (addAssessmentSection) {
    if (userRole === 'super_admin' || userRole === 'admin' || userRole === 'teacher') {
      addAssessmentSection.style.display = 'block';
    } else {
      addAssessmentSection.style.display = 'none';
    }
  }

  if (allAssessments.length === 0) {
    tbody.innerHTML = '';
    noAssessmentsMsg.style.display = 'block';
    return;
  }

  noAssessmentsMsg.style.display = 'none';

  tbody.innerHTML = allAssessments.map((assessment, index) => {
    // Determine status badge based on assessment state
    let statusBadge;
    if (assessment.isComplete === true) {
      // Processing complete and ready for students
      statusBadge = '<span style="background: #e6f4ea; color: #137333; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: 500;">✓ Ready</span>';
    } else if (assessment.chunkCount > 0) {
      // Processing in progress (text extracted, generating audio)
      statusBadge = '<span style="background: #fef7e0; color: #ea8600; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: 500;">⏳ Processing</span>';
    } else {
      // Newly uploaded, waiting for processing to start
      statusBadge = '<span style="background: #e8f0fe; color: #1967d2; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: 500;">📋 Queued</span>';
    }

    return `
      <tr style="border-bottom: 1px solid #e8eaed;" data-row-index="${assessment.rowIndex}">
        <td style="padding: 12px 8px;">${escapeHtml(assessment.fileName)}</td>
        <td style="padding: 12px 8px;">${statusBadge}</td>
        <td style="padding: 12px 8px;">
          <input type="text" value="${escapeHtml(assessment.className)}" data-field="className" style="width: 100%; padding: 4px 6px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px;">
        </td>
        <td style="padding: 12px 8px;">
          <input type="text" value="${escapeHtml(assessment.instructor)}" data-field="instructor" style="width: 100%; padding: 4px 6px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px;">
        </td>
        <td style="padding: 12px 8px;">
          <input type="text" value="${escapeHtml(assessment.password)}" data-field="password" style="width: 100%; padding: 4px 6px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px;">
        </td>
        <td style="padding: 12px 8px;">
          <textarea data-field="studentEmails" style="width: 100%; padding: 4px 6px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px; resize: vertical; min-height: 32px;">${escapeHtml(assessment.studentEmails)}</textarea>
        </td>
        <td style="padding: 12px 8px; text-align: center;">
          <button onclick="saveAssessmentRow(${assessment.rowIndex})" style="width: auto; padding: 6px 10px; font-size: 16px; margin-right: 4px; background: #34a853; color: white; border: none; border-radius: 4px; cursor: pointer;" title="Save">✓</button>
          ${userRole === 'super_admin' ? `
            <button onclick="deleteAssessmentRow(${assessment.rowIndex})" style="width: auto; padding: 6px 10px; font-size: 16px; margin-right: 4px; background: #d93025; color: white; border: none; border-radius: 4px; cursor: pointer;" title="Delete">✕</button>
            <button onclick="reprocessAssessmentRow(${assessment.rowIndex})" style="width: auto; padding: 6px 10px; font-size: 16px; background: #f8f9fa; color: #3c4043; border: 1px solid #dadce0; border-radius: 4px; cursor: pointer;" title="Re-run analysis and audio generation">⟳</button>
          ` : ''}
        </td>
      </tr>
    `;
  }).join('');
}

function addNewAssessment() {
  const addButton = document.getElementById('add-assessment-btn');
  const uploadMessage = document.getElementById('upload-message');
  const fileInput = document.getElementById('file-upload-input');
  const gdocInput = document.getElementById('gdoc-url-input');
  const className = document.getElementById('new-class-name').value.trim();
  const instructor = document.getElementById('new-instructor').value.trim();
  const password = document.getElementById('new-password').value.trim();
  const studentEmailsRaw = document.getElementById('new-student-emails').value;
  const studentEmails = parseEmailList(studentEmailsRaw);

  // Determine which tab is active
  const isFileTab = document.getElementById('upload-section-file').style.display !== 'none';

  uploadMessage.textContent = '';
  uploadMessage.style.color = '#5f6368';

  // Validate inputs
  if (isFileTab) {
    if (!fileInput.files || fileInput.files.length === 0) {
      uploadMessage.textContent = 'Please select a file to upload.';
      uploadMessage.style.color = '#d93025';
      return;
    }
  } else {
    if (!gdocInput.value.trim()) {
      uploadMessage.textContent = 'Please enter a Google Doc URL.';
      uploadMessage.style.color = '#d93025';
      return;
    }
  }

  addButton.disabled = true;
  addButton.textContent = 'Processing...';
  uploadMessage.textContent = 'Uploading file...';

  if (isFileTab) {
    // Handle file upload
    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
      const base64Data = e.target.result.split(',')[1]; // Remove data URL prefix
      const mimeType = file.type;

      uploadMessage.textContent = 'Uploading to Drive...';

      google.script.run
        .withSuccessHandler((result) => {
          if (result.error) {
            uploadMessage.textContent = 'Error: ' + result.error;
            uploadMessage.style.color = '#d93025';
            addButton.disabled = false;
            addButton.textContent = 'Add Assessment & Process';
            return;
          }

          uploadMessage.textContent = 'File uploaded. Adding to database...';

          // Add to spreadsheet
          addAssessmentToDatabase(result.fileUrl, className, instructor, password, studentEmails);
        })
        .withFailureHandler((error) => {
          uploadMessage.textContent = 'Upload failed: ' + error.message;
          uploadMessage.style.color = '#d93025';
          addButton.disabled = false;
          addButton.textContent = 'Add Assessment & Process';
        })
        .uploadAssessmentFile(adminSessionToken, file.name, base64Data, mimeType);
    };

    reader.onerror = function() {
      uploadMessage.textContent = 'Failed to read file.';
      uploadMessage.style.color = '#d93025';
      addButton.disabled = false;
      addButton.textContent = 'Add Assessment & Process';
    };

    reader.readAsDataURL(file);

  } else {
    // Handle Google Doc URL
    const docUrl = gdocInput.value.trim();
    uploadMessage.textContent = 'Processing Google Doc...';

    google.script.run
      .withSuccessHandler((result) => {
        if (result.error) {
          uploadMessage.textContent = 'Error: ' + result.error;
          uploadMessage.style.color = '#d93025';
          addButton.disabled = false;
          addButton.textContent = 'Add Assessment & Process';
          return;
        }

        if (result.message) {
          uploadMessage.textContent = result.message;
        }

        // Add to spreadsheet
        addAssessmentToDatabase(result.fileUrl, className, instructor, password, studentEmails);
      })
      .withFailureHandler((error) => {
        uploadMessage.textContent = 'Failed: ' + error.message;
        uploadMessage.style.color = '#d93025';
        addButton.disabled = false;
        addButton.textContent = 'Add Assessment & Process';
      })
      .handleGoogleDocUrl(adminSessionToken, docUrl);
  }
}

function addAssessmentToDatabase(fileUrl, className, instructor, password, studentEmails) {
  const addButton = document.getElementById('add-assessment-btn');
  const uploadMessage = document.getElementById('upload-message');

  google.script.run
    .withSuccessHandler((result) => {
      if (result.error) {
        uploadMessage.textContent = 'Error: ' + result.error;
        uploadMessage.style.color = '#d93025';
        addButton.disabled = false;
        addButton.textContent = 'Add Assessment & Process';
        return;
      }

      uploadMessage.textContent = result.message || 'Assessment added successfully!';
      uploadMessage.style.color = '#137333';

      // Reset form
      document.getElementById('file-upload-input').value = '';
      document.getElementById('gdoc-url-input').value = '';
      document.getElementById('new-class-name').value = '';
      document.getElementById('new-instructor').value = '';
      document.getElementById('new-password').value = '';
      document.getElementById('new-student-emails').value = '';

      addButton.disabled = false;
      addButton.textContent = 'Add Assessment & Process';

      // Reload assessments
      setTimeout(() => {
        uploadMessage.textContent = '';
        loadAllAssessments();
      }, 2000);
    })
    .withFailureHandler((error) => {
      uploadMessage.textContent = 'Failed to add: ' + error.message;
      uploadMessage.style.color = '#d93025';
      addButton.disabled = false;
      addButton.textContent = 'Add Assessment & Process';
    })
    .addNewAssessment(adminSessionToken, fileUrl, {
      className: className,
      instructor: instructor,
      password: password,
      studentEmails: studentEmails
    });
}

function saveAssessmentRow(rowIndex) {
  const row = document.querySelector(`tr[data-row-index="${rowIndex}"]`);
  if (!row) return;

  const className = row.querySelector('[data-field="className"]').value.trim();
  const instructor = row.querySelector('[data-field="instructor"]').value.trim();
  const password = row.querySelector('[data-field="password"]').value.trim();
  const studentEmailsRaw = row.querySelector('[data-field="studentEmails"]').value;
  const studentEmails = parseEmailList(studentEmailsRaw);

  console.log('Saving assessment row', rowIndex);

  google.script.run
    .withSuccessHandler((result) => {
      if (result.error) {
        alert('Error: ' + result.error);
        return;
      }
      alert('Assessment updated successfully!');
      loadAllAssessments();
    })
    .withFailureHandler((error) => {
      alert('Failed to update: ' + error.message);
    })
    .updateAssessmentRow(adminSessionToken, rowIndex, {
      className: className,
      instructor: instructor,
      password: password,
      studentEmails: studentEmails
    });
}

function deleteAssessmentRow(rowIndex) {
  if (!confirm('Are you sure you want to delete this assessment? This cannot be undone.')) {
    return;
  }

  console.log('Deleting assessment row', rowIndex);

  google.script.run
    .withSuccessHandler((result) => {
      if (result.error) {
        alert('Error: ' + result.error);
        return;
      }
      alert('Assessment deleted successfully!');
      loadAllAssessments();
    })
    .withFailureHandler((error) => {
      alert('Failed to delete: ' + error.message);
    })
    .deleteAssessmentRow(adminSessionToken, rowIndex);
}

function reprocessAssessmentRow(rowIndex) {
  if (!confirm('Re-process this assessment? This will clear existing audio data and regenerate it.')) {
    return;
  }

  console.log('Reprocessing assessment row', rowIndex);

  google.script.run
    .withSuccessHandler((result) => {
      if (result.error) {
        alert('Error: ' + result.error);
        return;
      }
      alert(result.message || 'Assessment reprocessing started!');
      loadAllAssessments();
    })
    .withFailureHandler((error) => {
      alert('Failed to reprocess: ' + error.message);
    })
    .reprocessAssessment(adminSessionToken, rowIndex);
}
