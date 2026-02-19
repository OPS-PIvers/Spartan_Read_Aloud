import os
import sys
import re
from playwright.sync_api import sync_playwright

# Read teacher.html
with open('teacher.html', 'r') as f:
    content = f.read()

# Mocking scriptlets
content = content.replace('<?!= sessionToken ?>', '"mock-session-token"')
content = content.replace('<?!= JSON.stringify(user) ?>', '{"email": "admin@example.com", "name": "Admin User", "userType": "super_admin"}')

# Remove any remaining <?!= ... ?>
content = re.sub(r'<\?!=.*?\?>', '', content)

# Mocking google.script.run
mock_google_script = """
<script>
window.google = {
  script: {
    run: {
      withSuccessHandler: function() { return this; },
      withFailureHandler: function() { return this; },
      getAllAssessments: function() { return this; },
      deleteAssessmentRow: function() {},
      saveAssessmentRow: function() {},
      reprocessAssessment: function() {}
    }
  }
};
</script>
"""
content = content.replace('<head>', '<head>' + mock_google_script)

# Injecting test data and calling renderTable
test_script = """
<script>
window.addEventListener('load', () => {
  const mainContent = document.getElementById('main-content');
  if (mainContent) mainContent.style.display = 'block';

  currentUserRole = 'super_admin';
  appConstants = {
    SUBMISSION_ADMIN_ONLY: false,
    SUBMISSION_ADMIN_ROLES: ['super_admin', 'admin']
  };

  assessmentList = [
    {
      rowIndex: 1,
      fileName: 'Test Assessment.pdf',
      className: 'Admin Class',
      instructor: 'Dr. Admin',
      password: 'admin-pass',
      accessExpires: '2025-12-31T23:59',
      readAloudEnabled: true,
      submissionEnabled: true,
      submissionDeliveryMode: 'bulk',
      submissionTimestamps: '2023-01-01',
      studentEmails: 'student1@example.com, student2@example.com',
      pdfUrl: 'https://example.com/test.pdf',
      status: 'ready'
    }
  ];

  if (typeof renderTable === 'function') {
    renderTable();
  }
});
</script>
"""
content = content.replace('</body>', test_script + '</body>')

# Save to mock file
with open('mock_teacher.html', 'w') as f:
    f.write(content)

# Take screenshot and verify computed styles
try:
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={'width': 1400, 'height': 800})
        file_path = os.path.abspath('mock_teacher.html')
        page.goto(f'file://{file_path}')

        page.wait_for_selector('#table-body tr', state='visible', timeout=10000)

        # Verify computed styles
        th_width = page.evaluate("window.getComputedStyle(document.querySelector('th:last-child')).minWidth")
        group_width = page.evaluate("window.getComputedStyle(document.querySelector('.action-group')).minWidth")
        btn_shrink = page.evaluate("window.getComputedStyle(document.querySelector('.action-btn')).flexShrink")

        print(f"TH Min-Width: {th_width}")
        print(f"Action Group Min-Width: {group_width}")
        print(f"Action Button Flex-Shrink: {btn_shrink}")

        # Capture the whole card containing the table
        page.evaluate("document.querySelector('section.card:last-of-type').scrollIntoView()")
        table_card = page.locator('section.card').last
        table_card.screenshot(path='fix_verification.png')

        browser.close()
    print("Verification complete.")
except Exception as e:
    print(f"Error: {e}")
