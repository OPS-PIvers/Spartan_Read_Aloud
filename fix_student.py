with open('student.html', 'r') as f:
    content = f.read()

old_onloaded = """function onAssessmentLoaded(result) {
  serverChunks = result.audioChunks;
  totalChunks = serverChunks.length;
  sessionToken = result.sessionToken;

  renderHtmlAssessment(result.assessmentHtml);
  initializeAudioToolbar();
  setLineSpacing(currentLineSpacing); // Apply default line spacing
  setupEventListeners();

  if (chunksLoaded < totalChunks) {
    bulkPreloadAudio();
  }
}"""

new_onloaded = """function onAssessmentLoaded(result) {
  serverChunks = result.audioChunks || [];
  totalChunks = serverChunks.length;
  sessionToken = result.sessionToken;

  renderHtmlAssessment(result.assessmentHtml);

  if (result.readAloudEnabled !== false) {
    initializeAudioToolbar();
    if (chunksLoaded < totalChunks) {
      bulkPreloadAudio();
    }
  } else {
    // Hide audio toolbar
    document.getElementById('audio-toolbar').style.display = 'none';
  }

  setLineSpacing(currentLineSpacing); // Apply default line spacing
  setupEventListeners();
}"""

content = content.replace(old_onloaded, new_onloaded)

with open('student.html', 'w') as f:
    f.write(content)
