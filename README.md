# Field Force Performance Dashboard — Setup & Run Guide

এই guide-এ ৪টা জিনিস আছে (তুমি যেভাবে চেয়েছ):
1. Code বসানো (install)
2. Streamlit-এ Run করা
3. Code hide/protect করা (limitations সহ honest details)
4. প্রতিদিন ব্যবহার করার নিয়ম

---

## 1. প্রথমবার Setup (একবারই করতে হবে)

### Step 1 — Python install
Windows/Mac-এ [python.org](https://www.python.org/downloads/) থেকে **Python 3.10+** install করো।
Install করার সময় "Add Python to PATH" checkbox টা টিক দিও।

### Step 2 — Files গুলো একটা ফোল্ডারে রাখো
```
dashboard-project/
 ├── dashboard.py
 ├── requirements.txt
 ├── data.xlsx                        (তোমার monthly workbook — এই নামেই রাখতে হবে)
 ├── ACI.png                          (company logo — optional, dashboard-এ header-এ দেখাবে)
 └── ACI_Premio_Plastics_Ltd_.jpg     (optional, header-এ দেখাবে)
```
⚠️ **`data.xlsx`** ফাইলটা এই নামেই রাখতে হবে — app চালু হলে এটা **automatically load** হয়ে
যাবে, কাউকে আর manually upload করতে হবে না। যাকে link পাঠাবে, সে সরাসরি খুলেই পুরো
dashboard দেখতে পাবে।

এই দুইটা logo file না থাকলেও app ঠিকভাবে চলবে (শুধু header-এ logo দেখাবে না)।

### Step 3 — Terminal/Command Prompt খুলে ঐ ফোল্ডারে যাও
```bash
cd path/to/dashboard-project
```

### Step 4 — Virtual environment বানাও (recommended, optional কিন্তু ভালো practice)
```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# Mac/Linux:
source venv/bin/activate
```

### Step 5 — Libraries install করো
```bash
pip install -r requirements.txt
```
(যদি `anthropic` না লাগে — AI narrative feature use না করলে — বাদ দিতে পারো requirements.txt থেকে।)

---

## 1.5. পরের মাসের নতুন Data দিতে চাইলে

নতুন মাসের workbook আসলে, শুধু পুরনো `data.xlsx` ফাইলটা **নতুন ফাইল দিয়ে replace** করো
(same নাম `data.xlsx` রেখে) — GitHub-এ upload করা থাকলে সেখানেও replace করে দিও। App
আবার চালু/reboot হলে নতুন data নিয়েই খুলবে। কাউকে কিছু upload করতে বলা লাগবে না।

Sidebar-এ একটা "Replace data for this session (optional)" uploader-ও আছে — সেটা শুধু
সাময়িকভাবে অন্য একটা file দিয়ে preview দেখার জন্য (session বন্ধ হলে আবার `data.xlsx`-এ
ফিরে আসবে)।

---

## 2. App Run করা (প্রতিবার)

```bash
streamlit run dashboard.py
```

এটা automatically browser-এ খুলে যাবে (usually `http://localhost:8501`)।
Sidebar থেকে তোমার মাসিক workbook (`.xlsx`, একই sheet-structure — Productivity Perfromance,
Highest_Place Found Data, Image_URL, 03 month Sales performance, Raw) upload করো।

Bandho করতে চাইলে terminal-এ `Ctrl + C` চাপো।

---

## 3. Code hide/protect করা — সত্যিটা জেনে রাখো

তুমি চেয়েছ যাতে "keu code na pai" — এটা technically পুরোপুরি possible না যদি app টা
তোমার নিজের কম্পিউটারে বা কারো hand-এ চলে (কারণ `.py` file তো তাদের কাছেই থাকে)। কিন্তু
practical ভাবে যা করা যায়:

- ✅ **App-এর ভেতরের UI** থেকে already "View source", "Report a bug", "☰ menu", এবং
  "Deploy" বাটন — সব hide করা আছে (কোডে CSS দিয়ে)। যে ইউজার শুধু dashboard browse করবে,
  সে কোনো menu/source link দেখবে না।
- ✅ **Streamlit Community Cloud**-এ deploy করলে, GitHub repo-টা **Private** রাখো —
  তাহলে শুধু browser-এ app টাই visible হবে, GitHub code কেউ দেখতে পারবে না।
- ✅ যদি নিজের server/VPS-এ চালাও, শুধু `.py` file আর data অন্য কাউকে না দিলেই হলো —
  end-user শুধু browser link পাবে, terminal/file access পাবে না।
- ⚠️ যদি কাউকে পুরো ফোল্ডার/laptop access দিতেই হয়, `.py` file আসলে hide করা যায় না
  (Python interpret করেই চালাতে হয়)। এক্ষেত্রে `pyarmor` বা commercial obfuscator দিয়ে
  code কে encrypt করা যায়, কিন্তু সেটা 100% unbreakable না — শুধু casual viewing আটকায়।

**সংক্ষেপে:** Streamlit Cloud + Private repo, অথবা নিজের server-এ শুধু browser-link শেয়ার
করাটাই সবচেয়ে বাস্তবসম্মত "hide code" সমাধান।

---

## 4. Sidebar-এর Policy Settings (গুরুত্বপূর্ণ)

তোমার rule ("প্রতি ২০ মিনিট idle = effective visit, কিন্তু একই জায়গায় ১.৫ ঘণ্টা থাকলে
effective না") default হিসেবে বসানো আছে, sidebar থেকে adjust করা যাবে:

| Setting | Default | মানে কী |
|---|---|---|
| Daily effective-visit target | 20 | প্রতিদিনের target visit সংখ্যা |
| Minutes of idle = 1 credited visit | 20 | কত মিনিট idle হলে ১টা visit credit হবে |
| Flag as "not effective" if same spot ≥ (hr) | 1.0 | একই জায়গায় এত ঘণ্টার বেশি থাকলে সেটা flag/বাতিল হবে |
| Flag 0-distance travel legs | ON | distance 0 হলেও "Travelling" দেখানো entry গুলো wastage হিসেবে ধরা |

এই numbers গুলো তোমার actual company SOP অনুযায়ী মিলিয়ে নিও — এটা একটা reasonable
default interpretation, official policy না।

---

## 5. Tabs কী কী দেখাবে

1. **Overview** — দুই employee-র photo/card, 3-month sales, flagged non-effective hours, May vs June bar chart
2. **Month-wise Cards** — May আর June আলাদা আলাদা card-এ productivity metrics
3. **Day/Area Time Analysis** — day-wise time spend, highest single-place time, সবচেয়ে বেশি cover করা area
4. **Effective Visit Audit** — reported vs সত্যিকারের credited effective visits, target-এর সাথে gap, flagged episodes list
5. **Wastage & "Fishing" Movement** — কোন দিন/জায়গায় সবচেয়ে বেশি সময় নষ্ট হয়েছে
6. **GPS Map & History** — date বেছে নিলে সেই দিনের পুরো movement map + click-through history table
7. **AI Low-Achiever Report** — rule-based narrative + চাইলে Anthropic API key দিয়ে AI-generated richer report
8. **Raw Data** — verification-এর জন্য raw sheets

---

## 6. AI Report Feature (Optional)

Sidebar-এ "Generate AI-written comparison report" টিক দিয়ে নিজের Anthropic API key
বসালে, Tab 7-এ একটা বাটন আসবে যেটা চাপলে Claude দিয়ে richer narrative লেখা হবে।
Key কোথাও save হয় না, শুধু ঐ session-এ ব্যবহার হয়।

Key না থাকলেও Tab 7-এ automatic rule-based summary সবসময় দেখাবে।
