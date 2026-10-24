# Kitchen-Primos-Survey.

<!DOCTYPE html>
<html lang="tl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Crispy Tuna with Cheese Evaluation Survey - Excel View</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- SheetJS for Export to Excel -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f3f4f6;
        }
        .excel-grid {
            border-collapse: collapse;
            width: 100%;
            background-color: white;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }
        .excel-grid th, .excel-grid td {
            border: 1px solid #d1d5db;
        }
        .col-header {
            background-color: #e5e7eb;
            color: #374151;
            font-weight: 600;
            font-size: 11px;
            text-align: center;
            user-select: none;
        }
        .row-header {
            background-color: #e5e7eb;
            color: #374151;
            font-weight: 600;
            font-size: 11px;
            text-align: center;
            width: 40px;
            user-select: none;
        }
        .survey-header {
            background-color: #1F4E79;
            color: white;
            font-weight: bold;
        }
        .sub-header {
            background-color: #D9E1F2;
            color: #1F4E79;
            font-weight: 600;
        }
        .rating-cell {
            text-align: center;
            cursor: pointer;
            transition: all 0.15s ease;
        }
        .rating-cell:hover {
            background-color: #e0e7ff !important;
        }
        .rating-cell.selected {
            background-color: #2563eb !important;
            color: white;
            font-weight: bold;
        }
        @media print {
            .no-print {
                display: none !important;
            }
            body {
                background-color: white;
            }
            .excel-grid {
                box-shadow: none;
            }
        }
    </style>
</head>
<body class="p-4 md:p-8">

    <!-- Top Toolbar -->
    <div class="max-w-6xl mx-auto mb-4 bg-white p-4 rounded-lg shadow flex flex-wrap justify-between items-center gap-4 no-print">
        <div class="flex items-center gap-3">
            <div class="bg-green-700 text-white p-2 rounded-lg flex items-center justify-center">
                <i class="fa-solid fa-file-excel text-2xl"></i>
            </div>
            <div>
                <h1 class="font-bold text-gray-800 text-lg">Crispy Tuna Evaluation Survey</h1>
                <p class="text-xs text-gray-500">Microsoft Excel View & Interactive Worksheet</p>
            </div>
        </div>
        
        <div class="flex items-center gap-2">
            <button onclick="clearRatings()" class="px-3 py-2 text-sm bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-md border flex items-center gap-2 transition">
                <i class="fa-solid fa-rotate-left"></i> Clear
            </button>
            <button onclick="window.print()" class="px-3 py-2 text-sm bg-blue-50 hover:bg-blue-100 text-blue-700 rounded-md border border-blue-200 flex items-center gap-2 transition">
                <i class="fa-solid fa-print"></i> Print
            </button>
            <button onclick="exportToExcel()" class="px-4 py-2 text-sm bg-green-700 hover:bg-green-800 text-white rounded-md font-medium flex items-center gap-2 shadow transition">
                <i class="fa-solid fa-download"></i> Export .XLSX
            </button>
        </div>
    </div>

    <!-- Excel Sheet Wrapper -->
    <div class="max-w-6xl mx-auto bg-white rounded-lg shadow overflow-hidden">
        <!-- Excel Ribbon Title Bar -->
        <div class="bg-green-800 text-white px-4 py-1.5 text-xs flex justify-between items-center no-print">
            <div class="flex items-center gap-4">
                <span class="font-semibold"><i class="fa-solid fa-table"></i> Crispy_Tuna_Evaluation_Survey.xlsx</span>
                <span class="bg-green-700 px-2 py-0.5 rounded text-[10px]">Saved</span>
            </div>
            <span class="text-gray-200">100% Zoom</span>
        </div>

        <div class="overflow-x-auto">
            <table class="excel-grid text-sm" id="surveyTable">
                <thead>
                    <!-- Excel Column Letters -->
                    <tr class="col-header">
                        <th class="row-header"></th>
                        <th class="w-12">A</th>
                        <th class="text-left px-4 py-1">B</th>
                        <th class="w-12">C</th>
                        <th class="w-12">D</th>
                        <th class="w-12">E</th>
                        <th class="w-12">F</th>
                        <th class="w-12">G</th>
                    </tr>
                </thead>
                <tbody>
                    <!-- Row 1: Title Header -->
                    <tr>
                        <td class="row-header">1</td>
                        <td colspan="7" class="survey-header text-center py-3 text-lg tracking-wide uppercase">
                            Crispy Tuna with Cheese - Evaluation Survey
                        </td>
                    </tr>

                    <!-- Row 2: Subtitle / Instructions -->
                    <tr>
                        <td class="row-header">2</td>
                        <td colspan="7" class="bg-gray-50 text-gray-600 text-xs text-center py-2 italic border-b">
                            Panuto: Paki-markahan ang bawat kategorya ayon sa iyong karanasan sa pagtikim gamit ang lebel ng rating mula 1 (Pinakamababa) hanggang 5 (Pinakamataas).
                        </td>
                    </tr>

                    <!-- Row 3: Blank Separator -->
                    <tr class="h-2">
                        <td class="row-header">3</td>
                        <td colspan="7" class="bg-gray-100"></td>
                    </tr>

                    <!-- Row 4: Column Headers -->
                    <tr>
                        <td class="row-header">4</td>
                        <td class="sub-header text-center py-2 font-bold w-16">No.</td>
                        <td class="sub-header text-left px-4 py-2 font-bold">Mga Tanong (Questions)</td>
                        <td class="sub-header text-center font-bold w-12">1</td>
                        <td class="sub-header text-center font-bold w-12">2</td>
                        <td class="sub-header text-center font-bold w-12">3</td>
                        <td class="sub-header text-center font-bold w-12">4</td>
                        <td class="sub-header text-center font-bold w-12">5</td>
                    </tr>

                    <!-- Questions 1 - 10 -->
                    <!-- Question 1 -->
                    <tr class="hover:bg-gray-50">
                        <td class="row-header">5</td>
                        <td class="text-center font-semibold text-gray-500">1</td>
                        <td class="px-4 py-2.5 text-gray-800">
                            <span class="font-semibold">Kagandahan ng Hitsura at Amoy:</span> Gaano ka-katakam-takam ang amoy at ayos ng pagkain nang ihain ito?
                        </td>
                        <td class="rating-cell" onclick="selectRating(this, 1)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 2)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 3)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 4)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 5)">[ ]</td>
                    </tr>

                    <!-- Question 2 -->
                    <tr class="bg-gray-50/50 hover:bg-gray-50">
                        <td class="row-header">6</td>
                        <td class="text-center font-semibold text-gray-500">2</td>
                        <td class="px-4 py-2.5 text-gray-800">
                            <span class="font-semibold">Sarap ng Keso:</span> Gaano kasarap at katakam-takam ang melted/creamy cheese na ginamit?
                        </td>
                        <td class="rating-cell" onclick="selectRating(this, 1)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 2)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 3)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 4)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 5)">[ ]</td>
                    </tr>

                    <!-- Question 3 -->
                    <tr class="hover:bg-gray-50">
                        <td class="row-header">7</td>
                        <td class="text-center font-semibold text-gray-500">3</td>
                        <td class="px-4 py-2.5 text-gray-800">
                            <span class="font-semibold">Kabuuan ng Lasa (Overall Taste):</span> Gaano mo nagustuhan ang pangkalahatang lasa ng Crispy Tuna with Cheese?
                        </td>
                        <td class="rating-cell" onclick="selectRating(this, 1)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 2)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 3)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 4)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 5)">[ ]</td>
                    </tr>

                    <!-- Question 4 -->
                    <tr class="bg-gray-50/50 hover:bg-gray-50">
                        <td class="row-header">8</td>
                        <td class="text-center font-semibold text-gray-500">4</td>
                        <td class="px-4 py-2.5 text-gray-800">
                            <span class="font-semibold">Dami at Presyo:</span> Makatarungan at sulit ang dami ng serving para sa presyo nito.
                        </td>
                        <td class="rating-cell" onclick="selectRating(this, 1)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 2)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 3)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 4)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 5)">[ ]</td>
                    </tr>

                    <!-- Question 5 -->
                    <tr class="hover:bg-gray-50">
                        <td class="row-header">9</td>
                        <td class="text-center font-semibold text-gray-500">5</td>
                        <td class="px-4 py-2.5 text-gray-800">
                            <span class="font-semibold">Pagiging Meryenda:</span> Gaano ito kagandang kainin bilang simpleng meryenda o finger food?
                        </td>
                        <td class="rating-cell" onclick="selectRating(this, 1)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 2)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 3)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 4)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 5)">[ ]</td>
                    </tr>

                    <!-- Question 6 -->
                    <tr class="bg-gray-50/50 hover:bg-gray-50">
                        <td class="row-header">10</td>
                        <td class="text-center font-semibold text-gray-500">6</td>
                        <td class="px-4 py-2.5 text-gray-800">
                            <span class="font-semibold">Pag-uulit ng Order:</span> Gaano kalaki ang posibilidad na oorder ka ulit ng produktong ito?
                        </td>
                        <td class="rating-cell" onclick="selectRating(this, 1)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 2)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 3)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 4)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 5)">[ ]</td>
                    </tr>

                    <!-- Question 7 -->
                    <tr class="hover:bg-gray-50">
                        <td class="row-header">11</td>
                        <td class="text-center font-semibold text-gray-500">7</td>
                        <td class="px-4 py-2.5 text-gray-800">
                            <span class="font-semibold">Rekomendasyon:</span> Gaano mo ire-rekomenda ang produktong ito sa iyong pamilya at mga kaibigan?
                        </td>
                        <td class="rating-cell" onclick="selectRating(this, 1)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 2)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 3)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 4)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 5)">[ ]</td>
                    </tr>

                    <!-- Question 8 -->
                    <tr class="bg-gray-50/50 hover:bg-gray-50">
                        <td class="row-header">12</td>
                        <td class="text-center font-semibold text-gray-500">8</td>
                        <td class="px-4 py-2.5 text-gray-800">
                            <span class="font-semibold">Interes sa Ibang Baryasyon:</span> Gaano ka ka-interesadong sumubok ng maanghang o ibang uri?
                        </td>
                        <td class="rating-cell" onclick="selectRating(this, 1)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 2)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 3)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 4)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 5)">[ ]</td>
                    </tr>

                    <!-- Question 9 -->
                    <tr class="hover:bg-gray-50">
                        <td class="row-header">13</td>
                        <td class="text-center font-semibold text-gray-500">9</td>
                        <td class="px-4 py-2.5 text-gray-800">
                            <span class="font-semibold">Pagkakaiba sa Iba:</span> Kumpara sa ibang renyas o seafood snacks sa merkado, mas natatangi ba ang produktong ito?
                        </td>
                        <td class="rating-cell" onclick="selectRating(this, 1)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 2)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 3)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 4)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 5)">[ ]</td>
                    </tr>

                    <!-- Question 10 -->
                    <tr class="bg-gray-50/50 hover:bg-gray-50">
                        <td class="row-header">14</td>
                        <td class="text-center font-semibold text-gray-500">10</td>
                        <td class="px-4 py-2.5 text-gray-800">
                            <span class="font-semibold">Pangkalahatang Kasiyahan:</span> Gaano ka nasiyahan sa iyong buong karanasan sa pagtikim ng produktong ito?
                        </td>
                        <td class="rating-cell" onclick="selectRating(this, 1)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 2)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 3)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 4)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 5)">[ ]</td>
                    </tr>

                    <!-- Bonus Questions Header Row -->
                    <tr>
                        <td class="row-header">15</td>
                        <td colspan="7" class="bg-amber-100 text-amber-900 font-bold px-4 py-1.5 text-xs uppercase tracking-wider">
                            <i class="fa-solid fa-star text-amber-600 mr-1"></i> Bonus Questions
                        </td>
                    </tr>

                    <!-- Bonus 1 -->
                    <tr class="bg-amber-50/40 hover:bg-amber-50">
                        <td class="row-header">16</td>
                        <td class="text-center font-semibold text-amber-700">B1</td>
                        <td class="px-4 py-2.5 text-gray-800">
                            <span class="font-semibold">Lutong at Crunchiness:</span> Gaano kalutong ang pagkain sa bawat kagat?
                        </td>
                        <td class="rating-cell" onclick="selectRating(this, 1)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 2)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 3)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 4)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 5)">[ ]</td>
                    </tr>

                    <!-- Bonus 2 -->
                    <tr class="bg-amber-50/40 hover:bg-amber-50">
                        <td class="row-header">17</td>
                        <td class="text-center font-semibold text-amber-700">B2</td>
                        <td class="px-4 py-2.5 text-gray-800">
                            <span class="font-semibold">Kaakibat na Sawsawan (Dipping Sauce):</span> Gaano kabagay ang ibinigay na sawsawan sa lasa ng Crispy Tuna with Cheese?
                        </td>
                        <td class="rating-cell" onclick="selectRating(this, 1)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 2)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 3)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 4)">[ ]</td>
                        <td class="rating-cell" onclick="selectRating(this, 5)">[ ]</td>
                    </tr>

                    <!-- Feedback Section Header -->
                    <tr>
                        <td class="row-header">18</td>
                        <td colspan="7" class="sub-header px-4 py-1.5 font-bold text-xs uppercase">
                            Mga Komento o Karagdagang Puna (Comments & Suggestions)
                        </td>
                    </tr>

                    <!-- Feedback Text Input -->
                    <tr>
                        <td class="row-header">19</td>
                        <td colspan="7" class="p-2">
                            <textarea class="w-full h-24 p-2 text-sm border border-gray-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="Isulat dito ang iyong mga puna o mungkahi..."></textarea>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- Excel Status Bar -->
        <div class="bg-gray-100 text-gray-500 px-4 py-1 text-xs bo