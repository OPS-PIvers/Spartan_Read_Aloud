const emailInput = document.getElementById('email-input');
const passwordInput = document.getElementById('password-input');
const loadButton = document.getElementById('load-button');
const loginErrorDiv = document.getElementById('login-error');

function loadAssessment() {
  loginErrorDiv.textContent = '';
  if (!emailInput.value || !passwordInput.value) {
    loginErrorDiv.textContent = 'Please enter both your email and the password.';
    return;
  }

  loadButton.disabled = true;
  loadButton.textContent = 'Loading...';

  google.script.run
    .withSuccessHandler(onAuthenticationSuccess)
    .withFailureHandler(onAssessmentLoadError)
    .authenticateUser(emailInput.value, passwordInput.value);
}

function onAuthenticationSuccess(result) {
  console.log('Authentication result:', result);

  if (result.error) {
    onAssessmentLoadError(result.error);
    return;
  }

  // Route based on user type
  if (result.userType === 'teacher' || result.userType === 'admin' || result.userType === 'super_admin') {
    google.script.run.withSuccessHandler(function(html) {
      document.getElementById('app-container').innerHTML = html;
    }).getTeacherView(result.sessionToken, result);
  } else if (result.userType === 'student') {
    google.script.run.withSuccessHandler(function(html) {
      document.getElementById('app-container').innerHTML = html;
    }).getStudentViewContent(result, emailInput.value, passwordInput.value);
  } else {
    onAssessmentLoadError('Invalid authentication response');
  }
}

function onAssessmentLoadError(error) {
  console.error('Assessment load error:', error);
  loginErrorDiv.textContent = error;
  loadButton.disabled = false;
  loadButton.textContent = 'Login';
}

passwordInput.addEventListener('keypress', function(event) {
  if (event.key === 'Enter') {
    event.preventDefault();
    loadButton.click();
  }
});

emailInput.addEventListener('keypress', function(event) {
    if (event.key === 'Enter') {
        event.preventDefault();
        passwordInput.focus();
    }
});
