# Kurdish Digital Calendar (KDC) 

Help keep the Kurdish Digital Calendar project alive and free! Your donation supports the development and maintenance of this valuable resource for the Kurdish community. Every contribution, big or small, makes a difference. Thank you for your generosity!

<a href="https://www.paypal.com/donate/?business=4U9ZWTGCL4XDG&amount=5&no_recurring=0&item_name=Support+the+Kurdish+Digital+Calendar%21+Your+donation+helps+keep+this+project+alive+and+free.+Thank+you+for+your+generosity%21&currency_code=USD">
    <img src="https://img.shields.io/badge/Donate-PayPal-blue.svg" alt="Donate" style="width:130px;height:auto;">
</a>
</br>
</br>


Download and try it out [💾here](https://github.com/6ebeng/Kurdish-Digital-Calendar/releases/latest/download/KDC.Installer.x64.x86.exe).

[![GitHub release](https://img.shields.io/github/v/release/6ebeng/kurdish-digital-calendar.svg)](https://github.com/6ebeng/Kurdish-Digital-Calendar/releases) 
[![GitHub issues](https://img.shields.io/github/issues/6ebeng/kurdish-digital-calendar.svg)](https://github.com/6ebeng/kurdish-digital-calendar/issues) 
[![GitHub forks](https://img.shields.io/github/forks/6ebeng/kurdish-digital-calendar.svg)](https://github.com/6ebeng/kurdish-digital-calendar/network) 
[![GitHub license](https://img.shields.io/github/license/6ebeng/kurdish-digital-calendar.svg)](https://github.com/6ebeng/Kurdish-Digital-Calendar/blob/master/LICENSE) 
[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/6ebeng/kurdish-digital-calendar)

The Kurdish Digital Calendar (KDC) is a versatile and user-friendly add-in designed for Microsoft Office applications. It allows users to seamlessly insert and convert dates between Gregorian, Hijri, Umm Al-Qura, and Kurdish calendars. The add-in supports various Office applications, including Word, Excel, PowerPoint, Outlook, COM Library for Access, Visio, and MS Project, ensuring consistency and accuracy in date handling across documents.


## Supported Dialects
- Kurdish (Central)
- Kurdish (Northern)

## Calendar Types
### Primary Calendars
- KDC recognizes these primary calendars in Kurdish, Arabic, and English:
	- Gregorian
	- Hijri
	- Umm al-Qura

### Reverse Conversion Calendars**
- These calendars are used for reverse conversion based on the chosen primary calendar:
	- Gregorian (English)
	- Gregorian (Arabic)
	- Gregorian (Kurdish Central)
	- Gregorian (Kurdish Northern)
	- Hijri (English)
	- Hijri (Arabic)
	- Hijri (Kurdish Central)
	- Hijri (Kurdish Northern)
	- Umm al-Qura (English)
	- Umm al-Qura (Arabic)
	- Umm al-Qura (Kurdish Central)
	- Umm al-Qura (Kurdish Northern)
	- Kurdish (Central)
	- Kurdish (Northern)

## Features

### KD Calendar Tab

#### Update Dates
- **Update Dates**
  - Updates all date entries in the document to reflect changes in calendar settings or formats.

### Converter (Selected Date)

#### Calendar Selection
- **Calendar**
  - Select the type of calendar system (e.g., Gregorian, Hijri, Umm Al-Qura).

#### Conversion
- Displays the selected date.
- **Switch**
  - Switch between different calendar formats.
- **Format**
  - Select the target format for conversion.
- Displays the converted date.
- **Convert**
  - Convert the selected date to the Kurdish calendar date or vice versa.

### Insert Date

#### Automatic Update
- **Update Automatically**
  - Automatically updates the date field when the document is opened or refreshed.

#### Date Format
- **Format**
  - Select a date format to use when inserting dates into your document.

#### Insert Today's Date
- **Today**
  - Insert today's date into the selected field.

#### Date Picker
- **Choose**
  - Open a date picker to select and insert a specific date.

### COM Library Integration
 - Includes a COM Library Reference KDC for VBA use in Word, Excel and Access.
	- **KDC for VBA**
		- Provides a set of functions for converting dates between different calendar systems.
		- Syntax: `KDC.ConvertDate(Date, fromCalendar, toCalendar, fromFormat, toFormat, targetDialect, isAddSuffix)` Default Dialect is Kurdish Central.
			- Parameters
				- Date: The date to convert.
				- fromCalendar: The calendar system of the input date.
				- toCalendar: The calendar system of the output date.
				- fromFormat: The format of the input date.
				- toFormat: The format of the output date.
				- targetDialect: The dialect of Kurdish to use ("ckb" for Central or "ku" for Northern).
				- isAddSuffix: A boolean value indicating whether to add suffixes to the output date.
			- Example: `KDC.ConvertDate("01/01/2024", "Gregorian", "Kurdish", "dd/MM/yyyy", "dddd, dd MMMM, yyyy", "Kurdish (Central)", true)` or `KDC.ConvertDate("01/01/2024", "Gregorian", "Kurdish", "dd/MM/yyyy", "dddd, dd MMMM, yyyy", "Kurdish (Northern)", true)`
			- Result: 
				- Kurdish Central Dialect : "دووشەممە، 11 بەفرانبار، 2723ی كوردی"
				- Kurdish Northern Dialect : "Duşem, 11 Berfanbar, 2723 Kurdî"
			- Note: The function supports dates from 0002-01-01 to 9999-12-31.

### User-Defined Functions in MS Excel
 - Call custom functions like `ConvertNumberToKurdishText` and `ConvertDateToKurdish` directly from Excel cells.
	- **ConvertNumberToKurdishText** (Support both Kurdish Central and Kurdish Northern Dialects)
		- Converts a number to Kurdish text.
		- Syntax: `ConvertNumberToKurdishText(Number, langcode as Optional)` Default Dialect is Kurdish Central.
			- Parameters
				- Number: The number to convert to Kurdish text.
				- langcode (optional): The dialect of Kurdish to use ("ckb" for Central or "ku" for Northern).
		- Example: `ConvertNumberToKurdishText(123456789)` or `ConvertNumberToKurdishText(123456789, "ku")` or `ConvertNumberToKurdishText(123456789, "ckb")`
		- Result: 
			- Kurdish Central Dialect : "سەد و بیست و سێ ملیۆن و چوار سەد و پەنجا و شەش هەزار و حەوت سه‌د و هه‌شتاو نۆ"
			- Kurdish Northern Dialect : "sed û bîst û sê milyon û çar sed û pêncî û şeş hezar û heft sed û heştê û neh"
		- Note: The function supports numbers up to 999,999,999,999,999,999.

	- **ConvertDateToKurdish** (Support both Kurdish Central and Kurdish Northern Dialects)
    - Converts a date to the Kurdish calendar.
        - Syntax: `ConvertDateToKurdish(Date, targetDialect as Optional, fromFormat as Optional, toFormat as Optional, targetCalendar as Optional, isAddSuffix as Optional)` Default Dialect is Kurdish Central.
		    - Parameters
			    - Date: The date to convert to the Kurdish calendar.
			    - targetDialect (optional): The dialect of Kurdish to use ("ckb" for Central or "ku" for Northern).
			    - fromFormat (optional): The format of the input date.
			    - toFormat (optional): The format of the output date.
			    - targetCalendar (optional): The calendar system to use for conversion.
			    - isAddSuffix (optional): A boolean value indicating whether to add suffixes to the output date.
            - Example: `ConvertDateToKurdish("01/01/2024", "Kurdish (Central)", "dd/MM/yyyy", "dddd, dd MMMM, yyyy", "Gregorian", true)` or `ConvertDateToKurdish(01/01/2024", "Kurdish (Northern)", "dd/MM/yyyy", "dddd, dd MMMM, yyyy", "Gregorian", true)`
            - Result: 
				- Kurdish Central Dialect : "دووشەممە، 11 بەفرانبار، 2723ی كوردی"
				- Kurdish Northern Dialect : "Duşem, 11 Berfanbar, 2723 Kurdî"
            - Note: The function supports dates from 0002-01-01 to 9999-12-31.

## Installation

### Prerequisites

- Microsoft Office (Word, Excel, PowerPoint, Outlook, Visio, Project)
- .NET Framework 4.8 or higher

### Using the Installer

1. Download the latest release from the [GitHub releases page](https://github.com/6ebeng/Kurdish-Digital-Calendar/releases).
2. Run the installer and select the components you wish to install.
3. Follow the on-screen instructions to complete the installation.

## Contributing

To contribute to the Kurdish Digital Calendar project:

1. Fork the repository on GitHub.
2. Create a new branch for your feature or bug fix.
3. Commit your changes and push your branch to GitHub.
4. Open a pull request with a description of your changes.

For more information, see our [contributing guidelines](https://github.com/6ebeng/kurdish-digital-calendar/blob/master/CONTRIBUTING.md).

## License

The Kurdish Digital Calendar is licensed under the MIT License. See the [LICENSE](https://github.com/6ebeng/kurdish-digital-calendar/blob/master/LICENSE) file for details.

## Credits

Developed and maintained by Tishko Rasoul ([Rekbin Devs](https://github.com/Rekbin-Devs)). Thanks to all contributors and supporters.

## Support

For issues or questions, open an issue on the [GitHub repository](https://github.com/6ebeng/kurdish-digital-calendar/issues) or contact us at [rekbin.devs@gmail.com](mailto:rekbin.devs@gmail.com).