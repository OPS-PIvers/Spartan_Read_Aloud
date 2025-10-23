<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Audio Toolbar Snippet</title>
    <!-- 1. Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- 2. Google Fonts (Inter) -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <!-- 3. Google Material Icons -->
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
    <style>
        /* Basic styling for Inter font */
        body {
            font-family: 'Inter', sans-serif;
        }
        /* Custom styles for the progress bar thumb */
        input[type="range"]::-webkit-slider-thumb {
            -webkit-appearance: none;
            appearance: none;
            width: 14px;
            height: 14px;
            background: #4f46e5; /* indigo-600 */
            border-radius: 9999px;
            cursor: pointer;
            margin-top: -5px; /* Center the thumb on the track */
            border: 2px solid white;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.2);
        }
        input[type="range"]::-moz-range-thumb {
            width: 14px;
            height: 14px;
            background: #4f46e5;
            border-radius: 9999px;
            cursor: pointer;
            border: 2px solid white;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.2);
        }
        /* Custom styles for volume slider thumb */
        #volume-slider::-webkit-slider-thumb {
            margin-top: -3px; /* Center on the thinner track */
            width: 12px;
            height: 12px;
        }
        #volume-slider::-moz-range-thumb {
            width: 12px;
            height: 12px;
        }
        
        /* This class is just to add some space so the toolbar doesn't cover content */
        body {
            /* Add padding to the bottom of your page equal to the toolbar's height */
            padding-bottom: 96px; /* h-24 */
            /* Add a min-height to ensure there's content to scroll past */
            min-height: 200vh;
        }
    </style>
</head>
<body class="bg-gray-100 dark:bg-gray-900 text-gray-900 dark:text-gray-100 transition-colors duration-300">

    <!-- 
      This is just dummy content so you can see the toolbar 
      properly positioned over a page.
    -->
    <div class="max-w-3xl mx-auto p-6 md:p-12">
        <h1 class="text-3xl font-bold text-indigo-600 dark:text-indigo-400">Your Page Content</h1>
        <div class="prose prose-lg dark:prose-invert mt-8 space-y-4">
            <p>
                Your page content goes here. The audio toolbar will be fixed to the bottom of the viewport, floating above this content.
            </p>
            <p>
                I've added a `padding-bottom: 96px;` (the height of the toolbar) to the `body` tag in the style block. This is a common practice to ensure that the fixed toolbar doesn't permanently hide the last few lines of your page content when you scroll to the very bottom.
            </p>
            <p>I've also added the popout menus for "Speed" and "Question List". Try clicking the icons in the toolbar!</p>
        </div>
    </div>


    <!-- 
      Modern Audio Toolbar
    -->
    <div id="audio-toolbar" class="fixed bottom-0 left-0 right-0 bg-white/95 dark:bg-gray-800/95 backdrop-blur-md border-t border-gray-200 dark:border-gray-700 shadow-[0_-5px_20px_rgba(0,0,0,0.05)] dark:shadow-[0_-5px_20px_rgba(0,0,0,0.2)]">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex items-center justify-between h-24">

                <!-- 
                  ZONE 1: LEFT (Context)
                -->
                <div class="w-64 hidden md:block flex-shrink-0">
                    <div class="flex items-center space-x-3">
                        <!-- Mock Artwork -->
                        <div class="w-12 h-12 bg-gradient-to-r from-indigo-500 to-purple-500 rounded-lg flex-shrink-0"></div>
                        <div>
                            <!-- TODO: Populate with dynamic data -->
                            <div id="toolbar-title" class="text-sm font-semibold text-gray-900 dark:text-white truncate" title="Chapter 1: The Council of the Gods">
                                Chapter 1: The Council...
                            </div>
                            <div id="toolbar-subtitle" class="text-xs text-gray-500 dark:text-gray-400">The Odyssey</div>
                        </div>
                    </div>
                </div>

                <!-- 
                  ZONE 2: CENTER (Controls & Progress)
                -->
                <div class="flex-1 min-w-0 px-4">
                    <div class="flex flex-col items-center gap-2">
                        <!-- Playback Controls -->
                        <div class="flex items-center justify-center gap-3 sm:gap-4">
                            <!-- Speed Control -->
                            <button id="speed-btn" title="Playback Speed (1.0x)" class="control-btn p-2 text-gray-500 hover:text-gray-900 dark:text-gray-400 dark:hover:text-white rounded-full hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors">
                                <span class="material-icons text-xl">slow_motion_video</span>
                            </button>
                            <!-- Skip Back 10s -->
                            <button id="skip-back-btn" title="Skip Back 10s" class="control-btn p-2 text-gray-500 hover:text-gray-900 dark:text-gray-400 dark:hover:text-white rounded-full hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors">
                                <span class="material-icons text-2xl">replay_10</span>
                            </button>
                            <!-- Previous Chunk -->
                            <button id="prev-btn" title="Previous Chunk" class="control-btn p-2 text-gray-600 hover:text-gray-900 dark:text-gray-300 dark:hover:text-white rounded-full hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors">
                                <span class="material-icons text-3xl">skip_previous</span>
                            </button>
                            <!-- Play/Pause Button (Hero Button) -->
                            <button id="play-pause-btn" title="Play" class="w-12 h-12 sm:w-14 sm:h-14 rounded-full bg-indigo-600 hover:bg-indigo-700 text-white flex items-center justify-center shadow-lg transition-transform transform hover:scale-105">
                                <span id="play-icon" class="material-icons text-4xl">play_arrow</span>
                                <span id="pause-icon" class="material-icons text-4xl hidden">pause</span>
                                <span id="loading-spinner" class="material-icons text-4xl hidden animate-spin">refresh</span>
                            </button>
                            <!-- Next Chunk -->
                            <button id="next-btn" title="Next Chunk" class="control-btn p-2 text-gray-600 hover:text-gray-900 dark:text-gray-300 dark:hover:text-white rounded-full hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors">
                                <span class="material-icons text-3xl">skip_next</span>
                            </button>
                            <!-- Skip Forward 10s -->
                            <button id="skip-forward-btn" title="Skip Forward 10s" class="control-btn p-2 text-gray-500 hover:text-gray-900 dark:text-gray-400 dark:hover:text-white rounded-full hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors">
                                <span class="material-icons text-2xl">forward_10</span>
                            </button>
                            <!-- Focus Mode (Suggestion) -->
                            <button id="focus-btn" title="Focus Mode" class="control-btn p-2 text-gray-500 hover:text-gray-900 dark:text-gray-400 dark:hover:text-white rounded-full hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors">
                                <span class="material-icons text-xl">center_focus_strong</span>
                            </button>
                        </div>
                        
                        <!-- Progress Bar & Time -->
                        <div class="w-full flex items-center gap-2">
                            <span id="current-time" class="text-xs font-mono text-gray-500 dark:text-gray-400 w-10 text-right">0:00</span>
                            <div class="flex-1 relative">
                                <!-- Progress Bar Track -->
                                <input id="timeline-slider" type="range" min="0" max="100" value="0" class="w-full h-1.5 bg-gray-200 dark:bg-gray-700 rounded-full appearance-none cursor-pointer">
                                <!-- Filled part -->
                                <div id="progress-fill" class="absolute top-1/2 left-0 h-1.5 bg-indigo-500 rounded-full -translate-y-1/2 pointer-events-none" style="width: 0%;"></div>
                                
                                <!-- Chunk Markers -->
                                <div id="chunk-markers" class="absolute top-1/2 left-0 w-full h-1.5 -translate-y-1/2 pointer-events-none">
                                    <!-- Example Markers (populate dynamically) -->
                                    <div class="absolute w-0.5 h-1.5 bg-white dark:bg-gray-800" style="left: 15%;"></div>
                                    <div class="absolute w-0.5 h-1.5 bg-white dark:bg-gray-800" style="left: 40%;"></div>
                                    <div class="absolute w-0.5 h-1.5 bg-white dark:bg-gray-800" style="left: 75%;"></div>
                                </div>
                            </div>
                            <span id="total-time" class="text-xs font-mono text-gray-500 dark:text-gray-400 w-10 text-left">0:00</span>
                        </div>
                    </div>
                </div>

                <!-- 
                  ZONE 3: RIGHT (Tools)
                -->
                <div class="w-64 hidden md:flex items-center justify-end flex-shrink-0 space-x-4">
                    <!-- Chunk List -->
                    <button id="chunk-list-btn" title="Question List" class="control-btn p-2 text-gray-500 hover:text-gray-900 dark:text-gray-400 dark:hover:text-white rounded-full hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors">
                        <span class="material-icons">list</span>
                    </button>
                    <!-- Volume Control -->
                    <div class="flex items-center gap-2">
                        <span id="volume-icon" class="material-icons text-lg text-gray-500 dark:text-gray-400">volume_up</span>
                        <input id="volume-slider" type="range" min="0" max="100" value="80" class="w-24 h-1 bg-gray-200 dark:bg-gray-700 rounded-full appearance-none cursor-pointer">
                    </div>
                </div>

            </div>
        </div>
    </div>
    <!-- End of Toolbar -->


    <!-- 
      NEW: Popout Menus
      - Positioned absolutely, hidden by default (`hidden`).
      - They will be positioned relative to their buttons using JavaScript.
    -->

    <!-- Speed Control Popup -->
    <div id="speed-popup" class="hidden fixed bottom-0 left-0 mb-28 w-48 bg-white dark:bg-gray-800 rounded-lg shadow-xl border border-gray-200 dark:border-gray-700 p-3 z-50">
        <p class="text-sm font-medium text-center text-gray-700 dark:text-gray-200 mb-2">Playback Speed</p>
        <div class="grid grid-cols-3 gap-2">
            <button class="speed-preset text-sm font-medium text-gray-600 dark:text-gray-300 py-1.5 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700" data-speed="0.5">0.5x</button>
            <button class="speed-preset text-sm font-medium text-gray-600 dark:text-gray-300 py-1.5 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700" data-speed="0.75">0.75x</button>
            <button class="speed-preset text-sm font-medium text-gray-600 dark:text-gray-300 py-1.5 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700" data-speed="1.0">1.0x</button>
            <button class="speed-preset text-sm font-medium text-gray-600 dark:text-gray-300 py-1.5 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700" data-speed="1.25">1.25x</button>
            <button class="speed-preset text-sm font-medium text-gray-600 dark:text-gray-300 py-1.5 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700" data-speed="1.5">1.5x</button>
            <button class="speed-preset text-sm font-medium text-gray-600 dark:text-gray-300 py-1.5 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700" data-speed="2.0">2.0x</button>
        </div>
        <!-- This is a visual cue, a "tail" for the popup -->
        <div class="absolute left-1/2 -bottom-2 -translate-x-1/2 w-4 h-4 bg-white dark:bg-gray-800 border-b border-r border-gray-200 dark:border-gray-700 transform rotate-45"></div>
    </div>

    <!-- Chunk List Popup -->
    <div id="chunk-list-popup" class="hidden fixed bottom-0 right-0 mb-28 w-72 bg-white dark:bg-gray-800 rounded-lg shadow-xl border border-gray-200 dark:border-gray-700 p-3 z-50">
        <p class="text-sm font-medium text-center text-gray-700 dark:text-gray-200 mb-2">Question List</p>
        <div id="chunk-list-items" class="max-h-64 overflow-y-auto space-y-1">
            <!-- Items will be populated dynamically, here are examples -->
            <button class="chunk-item w-full text-left text-sm text-gray-700 dark:text-gray-200 p-2 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700 truncate">1. Tell me, O Muse, of the man...</button>
            <button class="chunk-item w-full text-left text-sm text-gray-700 dark:text-gray-200 p-2 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700 truncate bg-indigo-50 dark:bg-indigo-900/50 text-indigo-700 dark:text-indigo-300 font-medium">2. But even so he saved not his...</button>
            <button class="chunk-item w-full text-left text-sm text-gray-700 dark:text-gray-200 p-2 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700 truncate">3. Of these things, goddess...</button>
            <button class="chunk-item w-full text-left text-sm text-gray-700 dark:text-gray-200 p-2 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700 truncate">4. Now all the rest, as many as...</button>
            <button class="chunk-item w-full text-left text-sm text-gray-700 dark:text-gray-200 p-2 rounded-md hover:bg-gray-100 dark:hover:bg-gray-700 truncate">5. But when, as the seasons...</button>
        </div>
        <!-- This is a visual cue, a "tail" for the popup -->
        <div class="absolute right-6 -bottom-2 w-4 h-4 bg-white dark:bg-gray-800 border-b border-r border-gray-200 dark:border-gray-700 transform rotate-45"></div>
    </div>


    <!-- 
      NEW: JavaScript for Popups
      - This script handles showing/hiding the popups and positioning them.
    -->
    <script>
        document.addEventListener('DOMContentLoaded', () => {
            const speedBtn = document.getElementById('speed-btn');
            const speedPopup = document.getElementById('speed-popup');

            const chunkListBtn = document.getElementById('chunk-list-btn');
            const chunkListPopup = document.getElementById('chunk-list-popup');
            
            // Function to position a popup above a button
            function positionPopup(button, popup) {
                const btnRect = button.getBoundingClientRect();
                const popupRect = popup.getBoundingClientRect();
                
                // Center the popup horizontally with the button
                let left = btnRect.left + (btnRect.width / 2) - (popupRect.width / 2);
                
                // Position popup above the button (mb-28 is 7rem/112px, plus a bit)
                const top = btnRect.top - popupRect.height - 10; // 10px spacing
                
                // Prevent popup from going off-screen
                if (left < 10) left = 10;
                if (left + popupRect.width > window.innerWidth - 10) {
                    left = window.innerWidth - popupRect.width - 10;
                }

                popup.style.left = `${left}px`;
                popup.style.top = `${top}px`;
                // Use fixed positioning relative to viewport top
                popup.style.position = 'fixed'; 
                // Override Tailwind's 'bottom-0'
                popup.style.bottom = 'auto'; 
            }

            // --- Speed Popup Logic ---
            speedBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                const isHidden = speedPopup.classList.contains('hidden');
                // Hide other popups
                chunkListPopup.classList.add('hidden');
                // Toggle this popup
                speedPopup.classList.toggle('hidden', !isHidden);
                
                if (isHidden) {
                    // Position it just before showing
                    positionPopup(speedBtn, speedPopup);
                }
            });

            // --- Chunk List Popup Logic ---
            chunkListBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                const isHidden = chunkListPopup.classList.contains('hidden');
                // Hide other popups
                speedPopup.classList.add('hidden');
                // Toggle this popup
                chunkListPopup.classList.toggle('hidden', !isHidden);
                
                if (isHidden) {
                    // Position it just before showing
                    positionPopup(chunkListBtn, chunkListPopup);
                }
            });

            // --- Global Click to Close ---
            document.addEventListener('click', (e) => {
                // If the click is not inside a popup or on a toggle button, hide all popups
                if (!e.target.closest('#speed-popup') && !e.target.closest('#chunk-list-popup')) {
                    speedPopup.classList.add('hidden');
                    chunkListPopup.classList.add('hidden');
                }
            });

            // --- Highlight Active Speed (Example) ---
            const speedPresets = document.querySelectorAll('.speed-preset');
            let currentSpeed = 1.0; // Your app would have the real value
            
            function updateSpeedHighlight() {
                speedPresets.forEach(btn => {
                    const speed = parseFloat(btn.dataset.speed);
                    btn.classList.toggle('bg-indigo-50', speed === currentSpeed);
                    btn.classList.toggle('dark:bg-indigo-900/50', speed === currentSpeed);
                    btn.classList.toggle('text-indigo-700', speed === currentSpeed);
                    btn.classList.toggle('dark:text-indigo-300', speed === currentSpeed);
                    btn.classList.toggle('font-medium', speed === currentSpeed);
                });
                // Update the button title
                speedBtn.title = `Playback Speed (${currentSpeed}x)`;
            }
            
            speedPresets.forEach(btn => {
                btn.addEventListener('click', () => {
                    currentSpeed = parseFloat(btn.dataset.speed);
                    // In your app, you would also call: setPlaybackSpeed(currentSpeed);
                    console.log("Set speed to", currentSpeed);
                    updateSpeedHighlight();
                    speedPopup.classList.add('hidden'); // Close popup on selection
                });
            });
            
            updateSpeedHighlight(); // Set initial highlight
        });
    </script>

</body>
</html>

