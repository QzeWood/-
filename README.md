# 文本亂碼修復工具（Text/Excel/Word Decode & Fix + Translate GUI）

> 一鍵修復常見的「亂碼」文本，支援 TXT/CSV/MD/HTML/JSON/YAML、XLSX、DOCX，並可選擇轉為繁簡/英/日。內建文字進度條，處理過程一目了然。

---

## 為什麼做這個工具（背景）

上網下載資料或跨系統拷貝檔案，常會遇到**編碼不一致**導致的亂碼（mojibake）：
例如 UTF-8、GBK、Big5、Shift-JIS 互相誤用，或是存檔/讀檔時的預設編碼不同。
這個工具會**自動嘗試多種解碼**並**逆轉常見 mojibake**，選擇最可讀的結果後輸出為 UTF-8，同時提供**可選翻譯**（繁簡/英/日）。

---

## 特色

* 🧠 **多引擎自動判斷**：多套編碼候選 + 逆轉常見 mojibake 路徑，挑選最優結果
* 📝 **多格式支援**：文字檔（.txt/.csv/.md/.json/.html…）、Excel（.xlsx）、Word（.docx）
* 🔤 **翻譯選擇**：不翻譯 / 簡中 / 繁中 / 英 / 日
* 🧱 **安全輸出**：不覆蓋原檔，另存為 `原檔名_fixed.*`（必要時自動加序號）
* 📊 **文字進度條**：在視窗最上方以 ASCII 條即時顯示 0–100%
* 🖥️ **簡潔 GUI**：無需命令列，點選就能用

---

## 下載（最簡單）

**懶人包：直接下載 .exe 即可使用（免安裝、免依賴）**
👉 若你怕麻煩、不想裝 Python 或任何套件，**建議直接前往 **Releases** 頁面下載：`文本亂碼修復工具.exe`

---

## 如何使用

1. 開啟程式（或執行 `.py` 版本）
2. 在上方選擇 **來源語言**（通常用「自動」即可）與 **目標語言**（或「不翻譯」）
3. 點 **「選擇文件並處理」**，可同時選多個檔案
4. 觀察上方**進度條**；完成後，下方會輸出**摘要與明細**
5. 修復後的檔案會存放在**原資料夾**，檔名加上 `_fixed`

---

## 支援的副檔名

* **文字檔**：`.txt .csv .tsv .srt .ass .md .json .yaml .yml .log .ini .cfg .html .htm .xml`
* **Excel**：`.xlsx`（需 `openpyxl`）
* **Word**：`.docx`（需 `python-docx`）

---

## 從原始碼執行（可選）

> 若你不使用 .exe，也可用 Python 執行。

1. 安裝 Python 3.9+
2. 安裝必要套件（Excel/Word/翻譯功能為可選）：

   ```bash
   pip install openpyxl python-docx opencc-python-reimplemented deep-translator
   ```
3. 執行：

   ```bash
   python 文字亂碼修復工具exe版-v1.0.py
   ```

> 備註：
>
> * **不翻譯**情境下，不需要 `opencc`/`deep-translator`。
> * .exe 版本已把依賴打包好，一般不需另外安裝。

---

## 隱私與檔案安全

* 程式於**本機端**處理檔案，不會上傳內容。
* 不會覆寫原始檔，輸出於同資料夾、加上 `_fixed`。
* 請務必保留原始檔做備份。

---

## 已知小提醒

* 超大檔可能處理較久；請留意磁碟空間。
* Word/Excel 內容修復主要針對**文字**，不動到版面與非文字元件。

---

## 由 ChatGPT 協助完成

本工具在設計、修正與文件撰寫過程中，有**ChatGPT 的協助**（需求釐清、API/打包指引、UI 微調建議等）。
最終實作、測試與發佈由 **Mr. Q** 負責。

---

## 交流與貢獻

* 有建議或 Bug？歡迎開 **Issue**
* 想提功能？歡迎發 **Pull Request**
* 一起把這個小工具變得更好用 🙌

---

## 授權

* 原始碼採用 **MIT License**（見本倉庫 `LICENSE`）。
* 隨附的第三方套件依其各自授權條款使用（已整理於 `THIRD_PARTY_LICENSES.md`，如有需要請一併檢視）。

---

## 致謝

* Open-source 社群提供的優秀套件（`openpyxl`, `python-docx`, `opencc-python-reimplemented`, `deep-translator` …）
