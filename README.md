```html
<!DOCTYPE html>
<html lang="tl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Crispy Tuna with Cheese - Evaluation Survey & Dashboard</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        brand: {
                            50: '#fffbeb',
                            100: '#fef3c7',
                            500: '#f59e0b',
                            600: '#d97706',
                            700: '#b45309',
                            800: '#92400e',
                            900: '#78350f',
                        },
                        tuna: {
                            500: '#ef4444',
                            700: '#b91c1c',
                            900: '#7f1d1d'
                        },
                        slateDark: '#1e293b'
                    }
                }
            }
        }
    </script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS for Export to Excel -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
        body {
            font-family: 'Plus Jakarta Sans', sans-serif;
            background-color: #f8fafc;
        }
        .rating-btn {
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .rating-btn:active {
            transform: scale(0.92);
        }
    </style>
</head>
<body class="min-h-screen text-slate-800 flex flex-col">

    <!-- Top Navigation Bar -->
    <header class="bg-slateDark text-white sticky top-0 z-50 shadow-md">
        <div class="max-w-4xl mx-auto px-4 py-3 flex items-center justify-between">
            <div class="flex items-center gap-3">
                <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-brand-500 to-amber-300 flex items-center justify-center text-slateDark text-xl font-black shadow-lg">
                    <i class="fa-solid fa-utensils"></i>
                </div>
                <div>
                    <h1 class="font-extrabold text-base md:text-lg leading-tight tracking-tight text-amber-400">Crispy Tuna</h1>
                    <p class="text-xs text-slate-300 font-medium">Evaluation App</p>
                </div>
            </div>

            <!-- Tab Buttons -->
            <nav class="flex bg-slate-800 p-1 rounded-xl border border-slate-700">
                <button id="tabFormBtn" onclick="switchTab('form')" class="px-3 md:px-4 py-1.5 rounded-lg text-xs md:text-sm font-semibold transition-all duration-200 bg-brand-500 text-slateDark shadow-sm">
                    <i class="fa-solid fa-pen-to-square mr-1.5"></i> Survey
                </button>
                <button id="tabDashBtn" onclick="switchTab('dash')" class="px-3 md:px-4 py-1.5 rounded-lg text-xs md:text-sm font-semibold text-slate-300 hover:text-white transition-all duration-200 flex items-center gap-1.5">
                    <i class="fa-solid fa-chart-pie"></i> Results 
                    <span id="navCountBadge" class="bg-amber-500/20 text-amber-300 text-[10px] px-1.5 py-0.5 rounded-full font-bold">0</span>
                </button>
            </nav>
        </div>
    </header>

    <!-- Main Container -->
    <main class="flex-grow max-w-4xl w-full mx-auto p-4 md:p-6">

        <!-- VIEW 1: SURVEY FORM -->
        <section id="formView" class="space-y-6">
            <!-- Header Card -->
            <div class="bg-gradient-to-r from-amber-500 via-amber-600 to-amber-700 text-white rounded-2xl p-6 shadow-xl relative overflow-hidden">
                <div class="absolute -right-6 -bottom-6 text-white/10 text-9xl pointer-events-none">
                    <i class="fa-solid fa-cheese"></i>
                </div>
                <div class="relative z-10">
                    <span class="bg-white/20 text-white text-xs font-bold px-3 py-1 rounded-full uppercase tracking-wider backdrop-blur-sm">Customer Feedback</span>
                    <h2 class="text-2xl md:text-3xl font-extrabold mt-2">Crispy Tuna with Cheese</h2>
                    <p class="text-amber-100 text-sm mt-1 max-w-xl">Paki-markahan ang bawat kategorya ayon sa iyong karanasan sa pagtikim gamit ang lebel ng rating mula 1 (Mababa) hanggang 5 (Pinakamataas).</p>
                </div>
            </div>

            <!-- Respondent Info -->
            <div class="bg-white rounded-2xl p-5 shadow-sm border border-slate-100">
                <label class="block text-xs font-bold uppercase text-slate-500 tracking-wider mb-2">Pangalan (Opsyonal)</label>
                <div class="relative">
                    <i class="fa-solid fa-user absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400"></i>
                    <input type="text" id="respondentName" placeholder="Pangalan (iwanang bakante kung gustong manatiling Anonymous)" class="w-full pl-10 pr-4 py-2.5 text-sm rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-brand-500 focus:border-brand-500">
                </div>
            </div>

            <!-- Survey Form -->
            <form id="surveyForm" onsubmit="submitSurvey(event)" class="space-y-4">
                
                <!-- Main Questions Container -->
                <div id="questionsContainer" class="space-y-4">
                    <!-- Dynamic Questions Injected Here -->
                </div>

                <!-- Comments & Suggestions -->
                <div class="bg-white rounded-2xl p-5 shadow-sm border border-slate-100 space-y-2">
                    <label class="block text-sm font-bold text-slate-800">
                        <i class="fa-solid fa-comment-dots text-brand-600 mr-1.5"></i> Mga Komento o Karagdagang Puna
                    </label>
                    <textarea id="comments" rows="3" placeholder="Isulat dito ang iyong mga puna, mungkahi, o karanasan sa pagkain..." class="w-full p-3 text-sm rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-brand-500 focus:border-brand-500"></textarea>
                </div>

                <!-- Submit Button -->
                <button type="submit" class="w-full bg-gradient-to-r from-brand-500 to-amber-600 hover:from-brand-600 hover:to-amber-700 text-white font-extrabold text-base py-4 rounded-2xl shadow-lg hover:shadow-xl transition duration-200 active:scale-[0.99] flex items-center justify-center gap-2">
                    <i class="fa-solid fa-paper-plane"></i> I-submit ang Ebalwasyon
                </button>
            </form>
        </section>

        <!-- VIEW 2: DASHBOARD / RESULTS -->
        <section id="dashView" class="hidden space-y-6">
            
            <!-- Dashboard Stats Overview -->
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div class="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
                    <div class="w-12 h-12 rounded-xl bg-amber-100 text-amber-600 flex items-center justify-center text-xl font-bold">
                        <i class="fa-solid fa-users"></i>
                    </div>
                    <div>
                        <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider">Kabuuan ng Sumagot</p>
                        <h3 id="statTotalResponses" class="text-2xl font-black text-slate-800">0</h3>
                    </div>
                </div>

                <div class="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
                    <div class="w-12 h-12 rounded-xl bg-emerald-100 text-emerald-600 flex items-center justify-center text-xl font-bold">
                        <i class="fa-solid fa-star"></i>
                    </div>
                    <div>
                        <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider">Overall Satisfaction</p>
                        <h3 id="statOverallAverage" class="text-2xl font-black text-slate-800">0.0 <span class="text-xs text-slate-400 font-normal">/ 5.0</span></h3>
                    </div>
                </div>

                <div class="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
                    <div class="w-12 h-12 rounded-xl bg-sky-100 text-sky-600 flex items-center justify-center text-xl font-bold">
                        <i class="fa-solid fa-thumbs-up"></i>
                    </div>
                    <div>
                        <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider">Reorder Intention</p>
                        <h3 id="statReorderRate" class="text-2xl font-black text-slate-800">0.0 <span class="text-xs text-slate-400 font-normal">/ 5.0</span></h3>
                    </div>
                </div>
            </div>

            <!-- Action Toolbar -->
            <div class="bg-white p-4 rounded-2xl border border-slate-100 shadow-sm flex flex-wrap gap-3 items-center justify-between">
                <h3 class="font-bold text-slate-800 text-sm md:text-base flex items-center gap-2">
                    <i class="fa-solid fa-chart-bar text-brand-600"></i> Pangkalahatang Resulta ng Bawat Kategorya
                </h3>
                <div class="flex items-center gap-2 w-full sm:w-auto">
                    <button onclick="exportExcel()" class="flex-1 sm:flex-none px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold transition flex items-center justify-center gap-2 shadow-sm">
                        <i class="fa-solid fa-file-excel"></i> Export to Excel
                    </button>
                    <button onclick="clearAllData()" class="px-3 py-2 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-xl text-xs font-bold transition flex items-center justify-center gap-1.5 border border-rose-200">
                        <i class="fa-solid fa-trash-can"></i> Clear Data
                    </button>
                </div>
            </div>

            <!-- Question Breakdowns -->
            <div id="resultsList" class="space-y-3">
                <!-- Breakdown Bars Injected Here -->
            </div>

            <!-- Comments Feed -->
            <div class="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm space-y-4">
                <h3 class="font-bold text-slate-800 text-sm md:text-base flex items-center gap-2">
                    <i class="fa-solid fa-comments text-amber-500"></i> Recent Feedback & Comments
                </h3>
                <div id="commentsFeed" class="space-y-3 max-h-80 overflow-y-auto pr-1">
                    <!-- Comments injected here -->
                </div>
            </div>
        </section>

    </main>

    <!-- Custom Success Modal -->
    <div id="successModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4 opacity-0 pointer-events-none transition-all duration-300">
        <div class="bg-white rounded-3xl p-6 md:p-8 max-w-sm w-full text-center space-y-4 shadow-2xl transform scale-90 transition-all duration-300" id="modalBox">
            <div class="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center text-3xl mx-auto shadow-inner">
                <i class="fa-solid fa-check"></i>
            </div>
            <h3 class="text-xl font-extrabold text-slate-800">Maraming Salamat!</h3>
            <p class="text-slate-500 text-sm">Matagumpay na naitala ang iyong ebalwasyon para sa Crispy Tuna with Cheese.</p>
            <button onclick="closeModal()" class="w-full bg-slateDark hover:bg-slate-800 text-white font-bold py-3 rounded-xl text-sm transition">
                Mag-submit ng Bago
            </button>
        </div>
    </div>

    <!-- Application Script -->
    <script>
        // Questions Database
        const questionsList = [
            { id: 1, text: "Gaano ka-katakam-takam ang amoy at ayos ng pagkain nang ihain ito?", label: "Kagandahan ng Hitsura at Amoy", isBonus: false },
            { id: 2, text: "Gaano kasarap at katakam-takam ang melted/creamy cheese na ginamit?", label: "Sarap ng Keso", isBonus: false },
            { id: 3, text: "Gaano mo nagustuhan ang pangkalahatang lasa ng Crispy Tuna with Cheese?", label: "Kabuuan ng Lasa (Overall Taste)", isBonus: false },
            { id: 4, text: "Makatarungan at sulit ang dami ng serving para sa presyo nito.", label: "Dami at Presyo", isBonus: false },
            { id: 5, text: "Gaano ito kagandang kainin bilang simpleng meryenda o finger food?", label: "Pagiging Meryenda", isBonus: false },
            { id: 6, text: "Gaano kalaki ang posibilidad na oorder ka ulit ng produktong ito?", label: "Pag-uulit ng Order", isBonus: false },
            { id: 7, text: "Gaano mo ire-rekomenda ang produktong ito sa iyong pamilya at mga kaibigan?", label: "Rekomendasyon", isBonus: false },
            { id: 8, text: "Gaano ka ka-interesadong sumubok ng maanghang o ibang uri?", label: "Interes sa Ibang Baryasyon", isBonus: false },
            { id: 9, text: "Kumpara sa ibang renyas o seafood snacks sa merkado, mas natatangi ba ang produktong ito?", label: "Pagkakaiba sa Iba", isBonus: false },
            { id: 10, text: "Gaano ka nasiyahan sa iyong buong karanasan sa pagtikim ng produktong ito?", label: "Pangkalahatang Kasiyahan", isBonus: false },
            { id: 11, text: "Gaano kalutong ang pagkain sa bawat kagat?", label: "Bonus 1: Lutong at Crunchiness", isBonus: true },
            { id: 12, text: "Gaano kabagay ang ibinigay na sawsawan sa lasa ng Crispy Tuna with Cheese?", label: "Bonus 2: Kaakibat na Sawsawan", isBonus: true }
        ];

        // Current user choices state
        let currentAnswers = {};

        // Local Storage Helper
        function getStoredData() {
            try {
                return JSON.parse(localStorage.getItem('crispy_tuna_survey_data')) || [];
            } catch (e) {
                return [];
            }
        }

        function saveStoredData(data) {
            localStorage.setItem('crispy_tuna_survey_data', JSON.stringify(data));
        }

        // Initialize Form UI
        function renderQuestions() {
            const container = document.getElementById('questionsContainer');
            container.innerHTML = '';

            questionsList.forEach((q, idx) => {
                const card = document.createElement('div');
                card.className = `p-5 rounded-2xl border shadow-sm transition-all duration-200 ${
                    q.isBonus ? 'bg-amber-50/60 border-amber-200' : 'bg-white border-slate-100'
                }`;

                card.innerHTML = `
                    <div class="flex items-start justify-between gap-2 mb-3">
                        <div>
                            <div class="flex items-center gap-2">
                                <span class="px-2 py-0.5 rounded-md text-[11px] font-bold ${
                                    q.isBonus ? 'bg-amber-200 text-amber-900' : 'bg-slate-100 text-slate-600'
                                }">
                                    ${q.isBonus ? 'BONUS' : 'Q' + q.id}
                                </span>
                                <h4 class="font-bold text-sm text-slate-800">${q.label}</h4>
                            </div>
                            <p class="text-xs text-slate-500 mt-1">${q.text}</p>
                        </div>
                    </div>

                    <!-- Rating Radio Buttons (1-5) -->
                    <div class="grid grid-cols-5 gap-1.5 md:gap-3 mt-3">
                        ${[1, 2, 3, 4, 5].map(rating => `
                            <button type="button" 
                                onclick="setRating(${q.id}, ${rating})" 
                                id="btn_q${q.id}_r${rating}"
                                class="rating-btn py-2.5 rounded-xl border border-slate-200 text-xs md:text-sm font-bold text-slate-600 bg-slate-50 hover:bg-amber-100 hover:border-amber-300 flex flex-col items-center justify-center gap-0.5">
                                <span class="text-xs ${rating === 5 ? 'text-amber-500' : 'text-slate-400'}"><i class="fa-solid fa-star"></i></span>
                                <span>${rating}</span>
                            </button>
                        `).join('')}
                    </div>
                `;
                container.appendChild(card);
            });
        }

        // Handle Rating Selection
        function setRating(qId, value) {
            currentAnswers[qId] = value;

            // Clear active styling for this question
            for (let i = 1; i <= 5; i++) {
                const btn = document.getElementById(`btn_q${qId}_r${i}`);
                if (btn) {
                    btn.className = "rating-btn py-2.5 rounded-xl border border-slate-200 text-xs md:text-sm font-bold text-slate-600 bg-slate-50 hover:bg-amber-100 hover:border-amber-300 flex flex-col items-center justify-center gap-0.5";
                }
            }

            // Set active styling for selected
            const activeBtn = document.getElementById(`btn_q${qId}_r${value}`);
            if (activeBtn) {
                activeBtn.className = "rating-btn py-2.5 rounded-xl border-2 border-brand-500 bg-amber-500 text-white font-extrabold text-xs md:text-sm flex flex-col items-center justify-center gap-0.5 shadow-md scale-[1.02]";
            }
        }

        // Submit Form
        function submitSurvey(e) {
            e.preventDefault();

            // Validate if all questions answered
            const unanswered = questionsList.filter(q => !currentAnswers[q.id]);
            if (unanswered.length > 0) {
                alert(`Paki-sagutan ang lahat ng katanungan. Kulang pa ng ${unanswered.length} na sagot.`);
                return;
            }

            const response = {
                id: Date.now(),
                date: new Date().toLocaleDateString('en-PH', { month: 'short', day: 'numeric', year: 'numeric', hour: '2-digit', minute: '2-digit' }),
                name: document.getElementById('respondentName').value.trim() || 'Anonymous',
                ratings: { ...currentAnswers },
                comments: document.getElementById('comments').value.trim()
            };

            const data = getStoredData();
            data.push(response);
            saveStoredData(data);

            showModal();
            updateDashboard();
        }

        // Switch View Tabs
        function switchTab(tab) {
            const formView = document.getElementById('formView');
            const dashView = document.getElementById('dashView');
            const tabFormBtn = document.getElementById('tabFormBtn');
            const tabDashBtn = document.getElementById('tabDashBtn');

            if (tab === 'form') {
                formView.classList.remove('hidden');
                dashView.classList.add('hidden');

                tabFormBtn.className = "px-3 md:px-4 py-1.5 rounded-lg text-xs md:text-sm font-semibold transition-all duration-200 bg-brand-500 text-slateDark shadow-sm";
                tabDashBtn.className = "px-3 md:px-4 py-1.5 rounded-lg text-xs md:text-sm font-semibold text-slate-300 hover:text-white transition-all duration-200 flex items-center gap-1.5";
            } else {
                formView.classList.add('hidden');
                dashView.classList.remove('hidden');

                tabDashBtn.className = "px-3 md:px-4 py-1.5 rounded-lg text-xs md:text-sm font-semibold transition-all duration-200 bg-brand-500 text-slateDa