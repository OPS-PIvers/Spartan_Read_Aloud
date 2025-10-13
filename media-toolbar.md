<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8"/>
<meta content="width=device-width, initial-scale=1.0" name="viewport"/>
<title>Assessment Reader</title>
<script src="https://cdn.tailwindcss.com?plugins=forms,typography,container-queries"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&amp;display=swap" rel="stylesheet"/>
<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet"/>
<script>
        tailwind.config = {
            darkMode: "class",
            theme: {
                extend: {
                    colors: {
                        primary: "#4F46E5",
                        "background-light": "#F9FAFB",
                        "background-dark": "#111827",
                        "card-light": "#FFFFFF",
                        "card-dark": "#1F2937",
                        "text-light": "#1F2937",
                        "text-dark": "#F9FAFB",
                        "subtle-light": "#6B7280",
                        "subtle-dark": "#9CA3AF",
                        "border-light": "#E5E7EB",
                        "border-dark": "#374151"
                    },
                    fontFamily: {
                        sans: ["Inter", "sans-serif"],
                    },
                    borderRadius: {
                        DEFAULT: "0.5rem",
                        lg: "0.75rem",
                        xl: "1rem",
                    },
                },
            },
        };
    </script>
<style>
        @keyframes pop-up {
            from {
                transform: translateY(100%) scale(0.95);
                opacity: 0;
            }
            to {
                transform: translateY(0) scale(1);
                opacity: 1;
            }
        }
        .popup-menu {
            animation: pop-up 0.3s ease-out forwards;
        }
    </style>
</head>
<body class="bg-background-light dark:bg-background-dark font-sans" x-data="{ speedMenuOpen: false, questionMenuOpen: false }">
<div class="min-h-screen flex items-center justify-center p-4">
<div class="w-full max-w-3xl bg-card-light dark:bg-card-dark rounded-xl shadow-lg p-8">
<header class="mb-8">
<h1 class="text-2xl font-bold text-center text-text-light dark:text-text-dark">Assessment Reader</h1>
</header>
<main class="prose prose-lg max-w-none text-text-light dark:text-text-dark prose-headings:text-text-light dark:prose-headings:text-text-dark">
<h2 class="text-xl font-semibold">2.3 Introduction to Memory</h2>
<ol class="list-decimal list-outside space-y-8">
<li>
<p>Devi spent time developing a set of note cards for an upcoming test that used word associations based on what the words meant in relation to each other. Which of the following did Devi use with this study method?</p>
<ul class="list-none pl-0 mt-4 space-y-2">
<li>(A) State-dependent memory</li>
<li>(B) Latent learning</li>
<li>(C) Effortful processing</li>
<li>(D) Procedural memory</li>
</ul>
</li>
<li>
<p>When Amy was seven years of age, she had a babysitter from France. During this time Amy learned to speak a little French. Years later, when Amy got to college, she signed up for a beginning French class. Amy learned the material in her French class much more quickly than her classmates did. Amy's rapid learning was most likely due to</p>
<ul class="list-none pl-0 mt-4 space-y-2">
<li>(A) implicit memory</li>
<li>(B) explicit memory</li>
<li>(C) savings or relearning</li>
<li>(D) flashbulb memory</li>
</ul>
</li>
<li>
<p>A researcher designs a study to test the effects of a new drug on memory. She has two groups of participants. One group receives the drug, and the other receives a placebo. The researcher then tests the participants' memory. This is an example of what kind of study?</p>
<ul class="list-none pl-0 mt-4 space-y-2">
<li>(A) Experimental</li>
<li>(B) Correlational</li>
<li>(C) Observational</li>
<li>(D) Case study</li>
</ul>
</li>
</ol>
</main>
</div>
</div>
<div class="fixed bottom-0 left-0 right-0 p-4 bg-transparent backdrop-blur-sm flex justify-center z-50">
<div class="w-full max-w-lg bg-card-light dark:bg-card-dark rounded-xl shadow-2xl border border-border-light dark:border-border-dark p-4">
<div class="flex items-center justify-between mb-2">
<div class="text-left">
<p class="text-sm font-semibold text-text-light dark:text-text-dark">2.3 Introduction to Memory</p>
<p class="text-xs text-subtle-light dark:text-subtle-dark">Chunk 6 of 11</p>
</div>
<div class="flex items-center space-x-2">
<button @click="speedMenuOpen = !speedMenuOpen" class="flex items-center justify-center text-subtle-light dark:text-subtle-dark hover:text-primary dark:hover:text-primary transition-colors h-10 w-10 rounded-full hover:bg-border-light dark:hover:bg-border-dark">
<span class="text-sm font-bold">1.25x</span>
</button>
<button @click="questionMenuOpen = !questionMenuOpen" class="relative text-subtle-light dark:text-subtle-dark hover:text-primary dark:hover:text-primary transition-colors h-10 w-10 rounded-full flex items-center justify-center hover:bg-border-light dark:hover:bg-border-dark">
<span class="material-icons">list_alt</span>
</button>
</div>
</div>
<div class="mb-3">
<div class="w-full bg-border-light dark:bg-border-dark rounded-full h-1.5">
<div class="bg-primary h-1.5 rounded-full" style="width: 45%"></div>
</div>
<div class="flex justify-between text-xs text-subtle-light dark:text-subtle-dark mt-1">
<span>01:02</span>
<span>-01:34</span>
</div>
</div>
<div class="flex items-center justify-center space-x-4">
<button class="text-subtle-light dark:text-subtle-dark hover:text-text-light dark:hover:text-text-dark transition-colors">
<span class="material-icons text-3xl">replay_10</span>
</button>
<button class="text-subtle-light dark:text-subtle-dark hover:text-text-light dark:hover:text-text-dark transition-colors">
<span class="material-icons text-4xl">skip_previous</span>
</button>
<button class="bg-primary text-white rounded-full p-4 shadow-lg transform hover:scale-105 transition-transform">
<span class="material-icons text-4xl">play_arrow</span>
</button>
<button class="text-subtle-light dark:text-subtle-dark hover:text-text-light dark:hover:text-text-dark transition-colors">
<span class="material-icons text-4xl">skip_next</span>
</button>
<button class="text-subtle-light dark:text-subtle-dark hover:text-text-light dark:hover:text-text-dark transition-colors">
<span class="material-icons text-3xl">forward_10</span>
</button>
</div>
</div>
<div @click.away="speedMenuOpen = false" class="popup-menu absolute bottom-full mb-3 w-full max-w-sm bg-card-light dark:bg-card-dark rounded-xl shadow-2xl border border-border-light dark:border-border-dark p-4" style="display: none;" x-show="speedMenuOpen" x-transition="">
<p class="text-center text-sm font-medium text-text-light dark:text-text-dark mb-3">Playback Speed</p>
<div class="grid grid-cols-3 gap-2">
<button class="px-3 py-2 text-sm rounded-lg text-subtle-light dark:text-subtle-dark hover:bg-border-light dark:hover:bg-border-dark transition-colors">0.5x</button>
<button class="px-3 py-2 text-sm rounded-lg text-subtle-light dark:text-subtle-dark hover:bg-border-light dark:hover:bg-border-dark transition-colors">0.75x</button>
<button class="px-3 py-2 text-sm rounded-lg text-subtle-light dark:text-subtle-dark hover:bg-border-light dark:hover:bg-border-dark transition-colors">1.0x</button>
<button class="px-3 py-2 text-sm rounded-lg bg-primary text-white transition-colors">1.25x</button>
<button class="px-3 py-2 text-sm rounded-lg text-subtle-light dark:text-subtle-dark hover:bg-border-light dark:hover:bg-border-dark transition-colors">1.5x</button>
<button class="px-3 py-2 text-sm rounded-lg text-subtle-light dark:text-subtle-dark hover:bg-border-light dark:hover:bg-border-dark transition-colors">2.0x</button>
</div>
</div>
<div @click.away="questionMenuOpen = false" class="popup-menu absolute bottom-full mb-3 w-full max-w-sm bg-card-light dark:bg-card-dark rounded-xl shadow-2xl border border-border-light dark:border-border-dark p-4 max-h-64 overflow-y-auto" style="display: none;" x-show="questionMenuOpen" x-transition="">
<p class="text-center text-sm font-medium text-text-light dark:text-text-dark mb-3">Questions</p>
<ul class="space-y-2">
<li><a class="block p-2 text-sm rounded-lg text-text-light dark:text-text-dark hover:bg-border-light dark:hover:bg-border-dark transition-colors truncate" href="#">1. Devi spent time developing a set of note cards...</a></li>
<li><a class="block p-2 text-sm rounded-lg text-text-light dark:text-text-dark hover:bg-border-light dark:hover:bg-border-dark transition-colors truncate" href="#">2. When Amy was seven years of age...</a></li>
<li><a class="block p-2 text-sm rounded-lg text-text-light dark:text-text-dark hover:bg-border-light dark:hover:bg-border-dark transition-colors truncate" href="#">3. A researcher designs a study to test the effects...</a></li>
<li><a class="block p-2 text-sm rounded-lg text-text-light dark:text-text-dark hover:bg-border-light dark:hover:bg-border-dark transition-colors truncate" href="#">4. Question four text would go here...</a></li>
<li><a class="block p-2 text-sm rounded-lg bg-primary/20 text-primary font-semibold transition-colors truncate" href="#">5. A group of ten students took a...</a></li>
<li><a class="block p-2 text-sm rounded-lg text-text-light dark:text-text-dark hover:bg-border-light dark:hover:bg-border-dark transition-colors truncate" href="#">6. The process of encoding refers to...</a></li>
</ul>
</div>
</div>
<script defer="" src="https://cdn.jsdelivr.net/gh/alpinejs/alpine@v2.x.x/dist/alpine.min.js"></script>

</body></html>