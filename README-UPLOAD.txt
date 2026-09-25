RAHMAN Question Generator - Option Shuffle Update
=================================================

কী যোগ করা হয়েছে
----------------
1. নতুন dropdown: "অপশন শাফেল করবেন? *"
   - No - অপশন শাফেল করবে না (Default)
   - Yes - অপশন শাফেল করবে

2. Yes নির্বাচন করলে প্রতিটি প্রশ্নের A/B/C/D option আলাদাভাবে random shuffle হবে।
3. সঠিক উত্তর (A/B/C/D) নতুন option position অনুযায়ী স্বয়ংক্রিয়ভাবে remap হবে।
4. Question preview, Question DOCX, Answer DOCX এবং Combined DOCX - সব জায়গায় একই remapped answer ব্যবহার হবে।
5. No নির্বাচন করলে আগের মতো option A/B/C/D মূল order-এ থাকবে; প্রশ্নের order shuffle আগের মতো চলবে।
6. History-তে option shuffle setting সংরক্ষণ হবে এবং Recreate/Edit করলে setting restore হবে।
7. Reset করলে option shuffle আবার No হবে।

GitHub-এ ব্যবহার
---------------
এই ZIP extract করে repository root-এ নিচের file-গুলো replace/upload করুন:
- index.html
- question-generator.html
- question-generator.css
- question-generator.js
- question-template.docx
- jszip.min.js

Repository-র অন্য file (যেমন style.css, profile.png, CV PDF) delete করবেন না।

পরীক্ষা
-------
Feature browser automation দিয়ে Yes/No দুই mode-এ পরীক্ষা করা হয়েছে।
Generated Combined DOCX-এ shuffled option-এর সঙ্গে answer letter সঠিকভাবে remap হয়েছে কিনা তাও যাচাই করা হয়েছে।
