
# MySheetsApp - Excel 名片自动同步助手 （中文版请看底部）
**Note**: This repository serves solely as the documentation and configuration guide for the internal tool CardSync-Bot. It does not contain source code.
---

# MySheetsApp - Excel Business Card Auto Sync Assistant (English)

## Introduction

MySheetsApp is a Windows desktop tool designed for batch importing and syncing Excel business card data to Google Sheets. It supports one-click import, automatic deduplication, note splitting, fax number normalization, and is suitable for users who need to organize and sync spreadsheet data to the cloud.

## Core Features
1. **One-click Excel Import**: Supports .xls/.xlsx files, automatically maps to Google Sheets after import.
2. **Smart Deduplication**: Automatically identifies duplicate data, supports multiple cases such as name+email, name without email, etc.
3. **Note Auto-Splitting**: The note field can be automatically split into multiple lines for further processing.
4. **Fax Number Normalization**: Automatically adds the prefix to fax numbers for consistent formatting.
5. **Tray Resident & Single Instance**: Only one window can be opened, tray icon is always visible, click to restore the main interface.
6. **Flexible Configuration**: You can customize Sheet ID, column index, folder path, etc. via config.json, no code changes required.

---

## Environment Setup & Requirements

For first-time use or when changing Google accounts, please complete the following Google API authorization steps, otherwise data cannot be synced.

### Step 1: Apply for Google API Access (First Time Only)
1. Log in to [Google Cloud Console](https://console.cloud.google.com/).
2. Click the project selector in the upper left, choose "New Project", name it `CardSyncBot`, and create.
3. In the left menu, click "APIs & Services" -> "Library".
4. Search for `Google Sheets API`, click in and select "Enable".

### Step 2: Create Service Account
1. In the left menu, click "APIs & Services" -> "Credentials".
2. Click "+ CREATE CREDENTIALS" at the top, select "Service Account".
3. **Name**: Enter `bot-user`, click "Create and Continue".
4. **Role**: **Very important!** Choose **Basic -> Editor**.
5. Click "Done" to finish.

### Step 3: Download Key (JSON File)
1. In the Credentials list, click the account email you just created (like `bot-user@...iam.gserviceaccount.com`).
2. Go to the "KEYS" tab, click "ADD KEY" -> "Create new key".
3. Choose **JSON** type, click Create.
4. The file will be downloaded automatically. **Rename it to `credentials.json`** and place it in the software root directory.

### Step 4: Authorize the Sheet
1. Open `bodhicardsynccredentials.json` with Notepad, copy the email address after "client_email".
2. Open your **Google Sheet (Master)**.
3. Click the **Share** button in the upper right.
4. Paste the service account email, set permission to **Editor**, and send.

---

## Quick Start

### 1. Download & Run

1. Get `MySheetsApp.exe` from the directory.
2. Double-click to run, no installation needed.

### 2. Excel File Preparation

- **Input Excel file column requirements:**

  | Col | Meaning        |
  |-----|---------------|
  | A   | Full Name     |
  | B   | Surname       |
  | C   | First Name    |
  | D   | Company       |
  | E   | Business Tel  |
  | F   | Business Email|
  | G   | Department    |
  | H   | Position      |
  | I   | Business Addr |
  | J   | Business Fax  |
  | K   | Mobile Phone  |
  | L   | Business URL  |
  | M   | Modified Time |
  | N   | Created Time  |
  | O   | Note          |

- **Note**: Please ensure the header order matches the table above, otherwise import may fail.

### 3. Start the Software

- Double-click `MySheetsApp.exe` to start.
- If already running, starting again will activate the open window.

### 4. Import Excel Files

1. Click the "Import" button, select Excel files to import (batch selection supported).
2. The software will automatically read, map fields, and deduplicate.

### 5. Field Mapping Rules (Automatic)

- The software automatically maps each input Excel column to the target Google Sheet column as follows:

  | Google Sheet Col | Source (Excel Col) | Description                                 |
  |-----------------|--------------------|---------------------------------------------|
  | A               | Blank              | For deduplication mark                      |
  | B               | Blank              | Categories                                  |
  | C               | Blank              | Type                                        |
  | D               | B                  | Surname                                     |
  | E               | C                  | First Name                                  |
  | F               | A                  | Full Name                                   |
  | G               | D                  | Company                                     |
  | H               | G                  | Department                                  |
  | I               | H                  | Position                                    |
  | J               | O (before ;)       | Industry (before semicolon in Note)         |
  | K               | Blank              | Function (reserved)                         |
  | L               | F                  | Email                                       |
  | M               | I                  | Address                                     |
  | N               | E                  | Business Tel                                |
  | O               | Blank              | Assistant name                              |
  | P               | Blank              | Assistant email                             |
  | Q               | Blank              | Assistant contact                           |
  | R               | K                  | Mobile Phone                                |
  | S               | O (after ;)        | Country (after semicolon in Note)           |
  | T               | J                  | Remark (Business Fax, auto add prefix Fax:) |
  | U               | N                  | Created Date                                |
  | V               | M                  | Modified Date                               |

- **Note (O col):**
  - If there is a semicolon, content before is "Industry", after is "Country/Note".
  - If no semicolon, all content is "Industry", country is left blank.

- **Fax (J col):**
  - Auto add prefix `Fax: `, will not duplicate if already present.

### 6. Deduplication & Marking

- The software automatically deduplicates by "Full Name + Email". Duplicates are marked in column A as "Duplicate (Confirmed) #number".

### 7. Sync to Google Sheets

1. On first use, a Google account authorization window will pop up. Please log in and authorize as prompted.
2. After sync, you can view the results in Google Sheets.

### 8. Tray & Single Instance

- When minimized or closed, the program stays in the system tray.
- Double-clicking the EXE or clicking the tray icon will bring up the window, no duplicate instances.

---

## Configuration

### 1. GUI Configuration (Recommended for Beginners)

Most common settings can be configured directly in the main interface:

- Open MySheetsApp.exe, click the "Settings" button in the upper right.
- In the popup, you can modify:
  - Google Sheet ID (SpreadsheetId)
  - Sheet Name (TargetSheetName)
  - Watch Folder Path (WatchDogFolder)
  - Excel Name/Email Column Index (InputNameIndex/InputEmailIndex)
  - Deduplication Mark Column (DuplicateMarkColumnLetter)
  - Auto Upload/Auto Start, etc.
- Click "Save" to apply changes.

**It is recommended to use the GUI for configuration, simple and intuitive for all users.**

---

### 2. config.json File (Advanced/Admin)

For batch deployment or advanced customization, you can edit `config.json` in the software directory:

```json
{
  "JsonKeyPath": "bodhicardsynccredentials.json", // Key file name
  "SpreadsheetId": "Your_Google_Sheet_ID_here", // Google Sheet ID
  "TargetSheetName": "Master", // Target sheet name
  "WatchDogFolder": "WatchDog", // Watch folder
  "InputNameIndex": 5, // Excel Name column index (A=0, B=1...)
  "InputEmailIndex": 11, // Excel Email column index
  "DuplicateMarkColumnLetter": "A", // Deduplication mark column in Google Sheet
  "EnableAutoStart": false, // Auto start
  "EnableAutoUpload": false // Auto upload
}
```

- **JsonKeyPath**: Key file name, must match the downloaded JSON file.
- **SpreadsheetId**: The string between `/d/` and `/edit` in the Google Sheet URL.
- **TargetSheetName**: Target sheet name, usually Master.
- **WatchDogFolder**: Watch folder name, put import files here.
- **InputNameIndex/InputEmailIndex**: Excel Name/Email column index (A=0, B=1...).
- **DuplicateMarkColumnLetter**: Which column in Google Sheet to write deduplication marks (e.g. A).
- **EnableAutoStart/EnableAutoUpload**: Auto start/upload (optional).

---

## FAQ

**Q1: Import prompts "File format incorrect"?**  
A: Please check if the Excel header order and column count match the requirements.

**Q2: Tray icon sometimes not visible?**  
A: The new version fixes this, tray icon is always visible. If still not, please restart the software.

**Q3: Sync to Google Sheets failed?**  
A: Please check your network and ensure Google account is authorized.

**Q4: How to exit the software?**  
A: Right-click the system tray icon and select "Exit" to fully close the program.

**Q5: Why is my Excel Name column the 6th, but config needs 5?**
**A:** The column index starts from 0.
* A=0
* B=1
* ...
* F=5
Please fill in `InputNameIndex` accordingly.

**Q6: What's the difference between "Duplicate (Confirmed)" and "Duplicate (Suspected)"?**
**A:**
* **Duplicate (Confirmed)**: Name and email both match, 100% same person.
* **Duplicate (Suspected)**: Name matches but one side lacks email, manual check recommended.

**Q7: How to switch to another Google Sheet?**
**A:** Just change SpreadsheetId in config.json, no code change needed.

---

### Support
If the program crashes or you encounter unsolvable bugs, please send the error logs from the logs folder to the developer.

---
**注意**：本仓库仅作为 `CardSync-Bot` 内部工具的操作文档与配置说明，不包含源代码。

## 简介

MySheetsApp 是一款 Windows 桌面工具，专为批量导入和同步 Excel 名片数据到 Google Sheets 设计。它支持一键导入、自动查重、备注拆分、传真号标准化等功能，适合需要整理和同步表格数据到云端的用户。

## 核心功能
1.  **一键导入 Excel**：支持 .xls/.xlsx 文件，导入后自动映射到 Google Sheets。
2.  **智能查重**：自动识别重复数据，支持姓名+邮箱、姓名缺邮箱等多种情况。
3.  **备注自动拆分**：备注字段可自动分割到多行，便于后续处理。
4.  **传真号标准化**：自动为传真号加前缀，格式统一。
5.  **托盘常驻与单实例**：程序只允许打开一个窗口，托盘图标始终可见，点击即可恢复主界面。
6.  **配置灵活**：通过 config.json 可自定义表格 ID、列索引、文件夹路径等，无需改代码。

---

## 环境配置与要求

首次使用或更换 Google 账号时，请务必完成以下 Google API 授权配置，否则无法同步数据。

### 步骤 1：申请 Google API 权限（仅首次）
1.  登录 [Google Cloud Console (谷歌云控制台)](https://console.cloud.google.com/)。
2.  点击左上角项目选择器，选择 "New Project" (新建项目)，命名为 `CardSyncBot`，点击创建。
3.  在左侧菜单点击 "APIs & Services" -> "Library" (库)。
4.  搜索 `Google Sheets API`，点击进入并选择 "Enable" (启用)。

### 步骤 2：创建服务账号 (Service Account)
1.  点击左侧菜单 "APIs & Services" -> "Credentials" (凭据)。
2.  点击顶部 "+ CREATE CREDENTIALS"，选择 "Service Account"。
3.  **Name**: 填写 `bot-user`，点击 "Create and Continue"。
4.  **Role (角色)**: **非常重要！** 选择 **Basic -> Editor (编辑者)**。
5.  点击 "Done" 完成。

### 步骤 3：下载密钥 (JSON 文件)
1.  在 Credentials 列表中，点击刚才创建的账号邮箱（类似 `bot-user@...iam.gserviceaccount.com`）。
2.  进入 "KEYS" 选项卡，点击 "ADD KEY" -> "Create new key"。
3.  选择 **JSON** 类型，点击 Create。
4.  文件会自动下载。**请将其重命名为 `credentials.json`**，并放入软件的根目录下。

### 步骤 4：给表格授权
1.  用记事本打开 `bodhicardsynccredentials.json`，复制 `"client_email"` 后面引号里的邮箱地址。
2.  打开您的 **Google Sheet (Master表)**。
3.  点击右上角 **Share (共享)** 按钮。
4.  粘贴机器人邮箱，权限设为 **Editor**，点击发送。

---

## 快速上手

### 1. 下载与安装

1. 从目录获取 `MySheetsApp.exe`。
2. 双击运行，无需安装。

### 2. Excel 文件准备

- **输入 Excel 文件格式要求**（每列含义）：

  | 列 | 含义         |
  |----|--------------|
  | A  | 全名         |
  | B  | 姓氏         |
  | C  | 名字         |
  | D  | 公司         |
  | E  | 商务电话     |
  | F  | 商务邮箱     |
  | G  | 部门         |
  | H  | 职称         |
  | I  | 商务地址     |
  | J  | 商务传真     |
  | K  | 手机电话     |
  | L  | 商务网址     |
  | M  | 修改时间     |
  | N  | 建立时间     |
  | O  | 附注         |

- **注意**：请确保表头顺序与上表一致，否则可能无法正确导入。

### 3. 启动软件

- 双击 `MySheetsApp.exe` 启动。
- 若已在运行，重复启动会自动激活已打开的窗口。

### 4. 导入 Excel 文件

1. 点击“导入”按钮，选择要导入的 Excel 文件（支持批量选择）。
2. 软件会自动读取文件内容，进行字段映射和查重处理。

### 5. 字段映射规则（自动完成，无需手动设置）

- 软件会自动将输入 Excel 的每一列映射到 Google Sheets 的目标列，具体如下：

  | Google Sheet 列 | 来源（Excel 列） | 说明                                   |
  |-----------------|------------------|----------------------------------------|
  | A               | 空白             | 用于查重标记                           |
  | B               | 空白             | Categories                             |
  | C               | 空白             | Type                                   |
  | D               | B                | Surname（姓氏）                        |
  | E               | C                | First Name（名字）                     |
  | F               | A                | Full Name（全名）                      |
  | G               | D                | Firm（公司）                           |
  | H               | G                | Department（部门）                     |
  | I               | H                | Position（职称）                       |
  | J               | O（分号前）      | Industry（附注分号前内容）             |
  | K               | 空白             | Function（预留）                       |
  | L               | F                | Email（商务邮箱）                      |
  | M               | I                | Address（商务地址）                    |
  | N               | E                | Contact（商务电话）                    |
  | O               | 空白             | Assistant name                         |
  | P               | 空白             | Assistant email                        |
  | Q               | 空白             | Assistant contact                      |
  | R               | K                | Contact（手机电话）                    |
  | S               | O（分号后）      | Country（附注分号后内容）              |
  | T               | J                | Remark（商务传真，自动加前缀 Fax: ）   |
  | U               | N                | Creation Date（建立时间）              |
  | V               | M                | Modified Date（修改时间）              |

- **附注（O列）说明**：  
  - 如果有分号，分号前为“行业”，分号后为“国家/地区及备注”；  
  - 如果没有分号，全部内容作为“行业”，国家/地区留空。

- **传真（J列）说明**：  
  - 自动加前缀 `Fax: `，如原内容已带前缀不会重复添加。

### 6. 查重与标记

- 软件会自动按“全名+邮箱”查重，重复项会在 Google Sheet 的 A 列标记“重复 (确) #编号”。

### 7. 同步到 Google Sheets

1. 首次使用会弹出 Google 账号授权窗口，请按提示登录并授权。
2. 同步完成后，可在 Google Sheets 查看导入结果。

### 8. 托盘与多开防护

- 最小化或关闭窗口时，程序会驻留在系统托盘。
- 再次双击 EXE 或点击托盘图标，窗口会自动弹出，不会重复打开多个实例。

---


## 软件配置

### 1. 图形化界面配置（推荐，小白专用）

大部分常用配置都可以直接在软件主界面完成，无需手动编辑文件：

- 打开 MySheetsApp.exe 后，点击右上角“设置”按钮。
- 在弹出的设置窗口中，可以修改：
  - Google 表格ID（SpreadsheetId）
  - Sheet名称（TargetSheetName）
  - 监控文件夹路径（WatchDogFolder）
  - Excel姓名/邮箱列索引（InputNameIndex/InputEmailIndex）
  - 查重标记列（DuplicateMarkColumnLetter）
  - 是否自动上传/开机自启等
- 修改后点击“保存”，配置会自动生效。

**建议优先使用图形化界面配置，简单直观，适合所有用户。**

---

### 2. config.json 文件配置（进阶/管理员专用）

如需批量部署或高级自定义，也可直接编辑软件目录下的 `config.json` 文件：

```json
{
  "JsonKeyPath": "bodhicardsynccredentials.json", // 密钥文件名
  "SpreadsheetId": "你的_Google_Sheet_ID_粘贴在这里", // Google 表格ID
  "TargetSheetName": "Master", // 目标表名
  "WatchDogFolder": "WatchDog", // 监控文件夹
  "InputNameIndex": 5, // Excel中“姓名”列索引（A=0，B=1...）
  "InputEmailIndex": 11, // Excel中“邮箱”列索引
  "DuplicateMarkColumnLetter": "A", // Google Sheet中查重标记写入的列
  "EnableAutoStart": false, // 是否开机自启
  "EnableAutoUpload": false // 是否自动上传
}
```

- **JsonKeyPath**: 密钥文件名，需与下载的 JSON 文件一致。
- **SpreadsheetId**: Google Sheet 网址中 `/d/` 和 `/edit` 之间的那串字符。
- **TargetSheetName**: 目标表名，通常为 Master。
- **WatchDogFolder**: 监控的文件夹名，导入文件请放这里。
- **InputNameIndex/InputEmailIndex**: Excel 中“姓名/邮箱”所在的列索引（A=0，B=1...）。
- **DuplicateMarkColumnLetter**: 查重结果写入 Google Sheet 的哪一列（如 A）。
- **EnableAutoStart/EnableAutoUpload**: 是否自动启动/自动上传（可选）。

---

## 常见问题

**Q1：导入时提示“文件格式不正确”？**  
A：请检查 Excel 文件的表头顺序和列数是否与要求一致。

**Q2：托盘图标有时不显示？**  
A：新版已修复，托盘图标始终可见，若仍有问题请重启软件。

**Q3：同步到 Google Sheets 失败？**  
A：请检查网络连接，并确保已正确授权 Google 账号。

**Q4：如何退出软件？**  
A：右键点击系统托盘图标，选择“退出”即可彻底关闭程序。

**Q5: 为什么我的 Excel 第 6 列是姓名，配置里却要写 5？**
**A:** 程序里的列索引是从 0 开始的。
* A列 = 0
* B列 = 1
* ...
* F列 = 5
请根据这个规则填写 `InputNameIndex`。

**Q6: “重复 (确)”和“重复 (暂)”有什么区别？**
**A:**
* **重复 (确)**：姓名和邮箱都一致，100%为同一人。
* **重复 (暂)**：姓名一致但有一方缺邮箱，建议人工核查。

**Q7: 我想换一个 Google 表格怎么办？**
**A:** 只需修改 config.json 里的 SpreadsheetId，无需改代码。

---

### 技术支持
如遇程序闪退或无法解决的 Bug，请将 logs 文件夹下的错误日志发送给开发者。
---
