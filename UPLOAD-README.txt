MD. ABDUR RAHMAN — Question Generator
======================================

এই প্যাকেজের সব ফাইল GitHub repository-র root folder-এ একই নাম রেখে Upload/Replace করুন।
বিশেষ করে question-generator.html, question-generator.css এবং question-generator.js তিনটি একসঙ্গে বদলাতে হবে।
question-template.docx ও jszip.min.js একই folder-এ থাকতে হবে।

কাজের নিয়ম
-----------
- Admin DOCX আপলোড করলে default Set A-D তৈরি হয়।
- উত্তর সেলে শুধু A, B, C অথবা D (ছোট হাতের অক্ষরও চলবে) থাকতে হবে।
- উত্তর ফাঁকা বা অন্য লেখা হলে সেই প্রশ্ন বাদ যাবে। পপআপ ও আপলোড ফর্মে প্রশ্নের নম্বর, DOCX সারি, প্রশ্ন এবং ভুল উত্তর দেখাবে।
- Admin DOCX সংশোধন করে আবার আপলোড করলে সেই প্রশ্ন জেনারেট করা যাবে।
- প্রতি সেটে প্রশ্নপত্র, উত্তরপত্র ও একত্রিত DOCX আলাদাভাবে ডাউনলোড করা যায়।
- হিস্ট্রি একই ব্রাউজারের IndexedDB-তে থাকে; অন্য ডিভাইস বা ব্রাউজারে তা দেখা যাবে না। ব্রাউজারের site data মুছলে হিস্ট্রিও মুছে যাবে।
- হিস্ট্রি থেকে মূল Admin DOCX, প্রতিটি আউটপুট এবং সব ফাইল ZIP আকারে আবার ডাউনলোড করা যায়।

Supported input columns:
(blank/serial optional) | Question | A | B | C | D | Answer | Explanation
