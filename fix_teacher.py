import sys

with open('teacher.html', 'r') as f:
    content = f.read()

# 1. Add checkbox to form
old_form = """          <div class="input-group">
            <label class="input-label">Student Emails (Comma separated)</label>
            <input type="text" id="meta-emails" class="form-control" placeholder="student1@school.com, student2@school.com">
          </div>
        </div>"""

new_form = """          <div class="input-group">
            <label class="input-label">Student Emails (Comma separated)</label>
            <input type="text" id="meta-emails" class="form-control" placeholder="student1@school.com, student2@school.com">
          </div>
          <div class="input-group" style="display: flex; align-items: center; gap: 10px; padding-top: 25px;">
            <input type="checkbox" id="meta-readaloud" checked style="width: 20px; height: 20px; cursor: pointer;">
            <label for="meta-readaloud" class="input-label" style="margin-bottom: 0; cursor: pointer;">Enable Read Aloud</label>
          </div>
        </div>"""

content = content.replace(old_form, new_form)

# 2. Update submitAssessment
old_submit = """      const meta = {
        className: document.getElementById('meta-class').value.trim(),
        instructor: document.getElementById('meta-instructor').value.trim(),
        password: document.getElementById('meta-password').value.trim(),
        studentEmails: document.getElementById('meta-emails').value.trim() // Server handles parsing
      };"""

new_submit = """      const meta = {
        className: document.getElementById('meta-class').value.trim(),
        instructor: document.getElementById('meta-instructor').value.trim(),
        password: document.getElementById('meta-password').value.trim(),
        studentEmails: document.getElementById('meta-emails').value.trim(), // Server handles parsing
        readAloudEnabled: document.getElementById('meta-readaloud').checked
      };"""

content = content.replace(old_submit, new_submit)

# 3. Update table headers
old_headers = """              <th style="width: 13%;">Password</th>
              <th style="width: 23%;">Students</th>"""

new_headers = """              <th style="width: 10%;">Password</th>
              <th style="width: 8%;">Read Aloud</th>
              <th style="width: 20%;">Students</th>"""

content = content.replace(old_headers, new_headers)

# 4. Update renderTable
old_render = """            <td><input class="table-edit-input" data-field="password" value="${escapeHtml(item.password)}" placeholder="--"></td>
            <td>
              <div class="email-chips-container\""""

new_render = """            <td><input class="table-edit-input" data-field="password" value="${escapeHtml(item.password)}" placeholder="--"></td>
            <td style="text-align:center;">
              <input type="checkbox" data-field="readAloudEnabled" ${item.readAloudEnabled !== false ? 'checked' : ''} style="width:18px; height:18px; cursor:pointer;">
            </td>
            <td>
              <div class="email-chips-container\""""

content = content.replace(old_render, new_render)

# 5. Update saveRow
old_save = """      const data = {
        className: row.querySelector('[data-field="className"]').value,
        instructor: row.querySelector('[data-field="instructor"]').value,
        password: row.querySelector('[data-field="password"]').value,
        studentEmails: emails.join(', ')
      };"""

new_save = """      const data = {
        className: row.querySelector('[data-field="className"]').value,
        instructor: row.querySelector('[data-field="instructor"]').value,
        password: row.querySelector('[data-field="password"]').value,
        studentEmails: emails.join(', '),
        readAloudEnabled: row.querySelector('[data-field="readAloudEnabled"]').checked
      };"""

content = content.replace(old_save, new_save)

with open('teacher.html', 'w') as f:
    f.write(content)
