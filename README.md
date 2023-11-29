# Automate-GoogleCalendarGmail-Out-Of-Office
An Apps Script to automate Gmail out-of-office responses based on Google Calendar events. This script checks for specific keywords in calendar events and activates out-of-office replies accordingly. Feel free to customize the response body and explore potential enhancements. For questions or collaboration, leave a comment.
---
# GoogleCalendar/Gmail Out-of-Office Automation

Automate Gmail's out-of-office replies based on Google Calendar events using Google Apps Script.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Example](#example)
- [Contributing](#contributing)
- [License](#license)
- [Acknowledgments](#acknowledgments)
- [Contact](#contact)
- [Troubleshooting](#troubleshooting)
- [Changelog](#changelog)

## Installation

1. Clone this repository.
2. Open the script project in Google Apps Script editor.
3. Enable the necessary APIs, including Gmail API and Calendar API.
4. Configure your `appsscript.json` manifest file (See [Editing a manifest](#editing-a-manifest)).
5. Save your changes.

## Usage

1. Run the `activateOutOfOffice` function in the script editor.
2. Ensure proper labeling of Google Calendar events with keywords (e.g., 'Out of Office', 'Vacation').
3. The script will automatically activate out-of-office replies based on the events in your calendar.

## Configuration

### Editing a manifest

1. Open the script project in the Apps Script editor.
2. Click Project Settings settings.
3. Select the Show "appsscript.json" manifest file in editor checkbox.
4. Edit the `appsscript.json` file directly in the editor and save any changes.
5. To hide the manifest file, repeat the steps and clear the Show "appsscript.json" manifest file in editor checkbox.

## Example

Consider the following Google Calendar event:

- **Event Name:** Out of Office
- **Start Date:** 2023-11-25
- **End Date:** 2023-11-30

The script will activate an out-of-office reply during this period.

## How to Implement

[Instruction](https://www.linkedin.com/pulse/automate-gmails-out-of-office-based-google-calendar-event-gyasi-nykoe?trk=public_post_feed-article-content)

## Contributing

Feel free to contribute to the project. Please follow our [Contribution Guidelines](CONTRIBUTING.md).

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.

## Acknowledgments

- Thanks to the Google Apps Script community for inspiration and support.

## Contact

For any questions or feedback, reach out to Reximus at fewzhon@gmail.com.

## Troubleshooting

If you encounter any issues, refer to the [Troubleshooting](#troubleshooting) section in the documentation.

## Changelog

### Version 1.0.0 (2023-11-30)

- Initial release.
