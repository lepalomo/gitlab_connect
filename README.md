# GitLab Connect

GitLab Connect is a Google Apps Script project that integrates Google Sheets with the GitLab API. It allows users to manage GitLab data directly from a spreadsheet, making it easier to get merge requests data.

## Features
- 📊 Fetch GitLab data and display it in a Google Spreadsheet.
- 📂 Organize GitLab data into three different sheets for better merge request data tracking and metrics.

## Installation
1. Open your Google Spreadsheet.
2. Navigate to **Extensions > Apps Script**.
3. Copy and paste the contents of this repository into the Apps Script editor.
4. Save and deploy the script.

## Configuration
1. 🔑 Set up your GitLab API credentials in the script.
2. 🌍 Define your GitLab instance URL and project ID.
3. 📑 Ensure your Google Spreadsheet has the necessary structure to receive and update GitLab data.

## Main Functions
### `manageGitLabData()`
- 🚀 The core function that interacts with the GitLab API.
- 📥 Fetches, processes, and updates data in Google Sheets.

### `write*` Functions
- ✍️ Any function starting with `write` is responsible for writing data to the spreadsheet.
- 📌 These functions handle different aspects of GitLab data and distribute them across three designated sheets.

## Usage
1. ▶️ Run `manageGitLabData()` to fetch and update GitLab information.
2. 📝 Use `write*` functions to modify or add GitLab-related data to the spreadsheet.
3. ⏰ Automate script execution using Google Apps Script triggers for periodic updates.

## Contributing
1. 🍴 Fork this repository.
2. 🌱 Create a new branch (`feature-branch-name`).
3. 💾 Commit your changes.
4. 📤 Push the branch and create a pull request.

## License
📝 This project is licensed under the MIT License.

## Contact
📬 For any issues or feature requests, open an issue on GitHub or reach out to the repository owner.

