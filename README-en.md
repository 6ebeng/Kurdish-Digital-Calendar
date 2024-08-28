<div align="center">
    <img src="https://github.com/6ebeng/Kurdish-Digital-Calendar/blob/master/Assets/Positive%20Logo%20Only.png" alt="KDC" style="width:200px;height:auto;">
</div>

# Kurdish Digital Calendar (KDC) 

Help keep the Kurdish Digital Calendar project alive and free! Your donation supports the development and maintenance of this valuable resource for the Kurdish community. Every contribution, big or small, makes a difference. Thank you for your generosity!

<a href="https://www.paypal.com/donate/?business=4U9ZWTGCL4XDG&amount=5&no_recurring=0&item_name=Support+the+Kurdish+Digital+Calendar%21+Your+donation+helps+keep+this+project+alive+and+free.+Thank+you+for+your+generosity%21&currency_code=USD">
<img src="https://img.shields.io/badge/Donate-PayPal-blue.svg" alt="Donate" style="width:130px;height:auto;">
</a>
<a href="#">
<img src="https://img.shields.io/badge/7507270392-FastPay-red.svg" alt="Donate" style="width:175px;height:auto;"></a>
</a>
<a href="#">
<img src="https://img.shields.io/badge/7507270392-FIB-cyan.svg" alt="Donate" style="width:140px;height:auto;">
</a>

Donate by Crypto USDT-TRC20: `TWtHokKWbRGG5R4BahUoggCMA8rL1TFttW`

</br>
</br>

Download and try it out [💾here](https://github.com/6ebeng/Kurdish-Digital-Calendar/releases/latest/download/KDC.Installer.exe).

[![GitHub release](https://img.shields.io/github/v/release/6ebeng/kurdish-digital-calendar.svg)](https://github.com/6ebeng/Kurdish-Digital-Calendar/releases) 
[![GitHub issues](https://img.shields.io/github/issues/6ebeng/kurdish-digital-calendar.svg)](https://github.com/6ebeng/kurdish-digital-calendar/issues) 
[![GitHub forks](https://img.shields.io/github/forks/6ebeng/kurdish-digital-calendar.svg)](https://github.com/6ebeng/kurdish-digital-calendar/network) 
[![GitHub license](https://img.shields.io/github/license/6ebeng/kurdish-digital-calendar.svg)](https://github.com/6ebeng/Kurdish-Digital-Calendar/blob/master/LICENSE) 
[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/6ebeng/kurdish-digital-calendar)


<a href="https://www.youtube.com/watch?v=gs6IB8x-jhY" target="_blank" rel="noopener noreferrer">
    <img src="https://raw.githubusercontent.com/6ebeng/Kurdish-Digital-Calendar/master/Assets/thumbnail_1.jpg" width="400" height="auto" />
</a>


</br>
</br>

The Kurdish Digital Calendar (KDC) is a versatile and user-friendly add-in designed for Microsoft Office applications. It allows users to seamlessly insert and convert dates between Gregorian, Hijri, Umm Al-Qura, and Kurdish calendars. The add-in supports various Office applications, including Word, Excel, PowerPoint, Outlook, COM Library for Access, Visio, and MS Project, ensuring consistency and accuracy in date handling across documents.

</br>
<img src="https://raw.githubusercontent.com/6ebeng/Kurdish-Digital-Calendar/master/Assets/Screenshots/Screenshot%20MS%20Word%202024-07-16%20010355.png" alt="Screenshot1" style="width:auto;height:auto;">

## Supported Dialects
- Kurdish (Central) - Sorani
- Kurdish (Northern) - Kurmanji

## Calendar Types
### Primary Calendars
- KDC recognizes these primary calendars in Kurdish, Arabic, and English:
	- Gregorian
	- Hijri
	- Umm al-Qura

### Reverse Conversion Calendars
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

## Supported Date Formats
- dddd, dd MMMM, yyyy
- dddd, dd/MM/yyyy
- dd MMMM, yyyy
- MMMM dd, yyyy
- dd/MM/yyyy
- MM/dd/yyyy
- yyyy/MM/dd
- MMMM yyyy
- MM/yyyy
- MMMM
- yyyy

## Icon themes
- Light
- Dark

## Features

### KD Calendar Tab

#### Settings
- **Settings**
  - Open the settings dialog to configure the calendar settings, such as selecting the dialect, icon theme, add suffix calendar name, and enabling/disabling automatic date updates upon load document or open document.

#### Update Dates
- **Update Dates**
  - Updates all date entries in the document to reflect changes in calendar settings or formats.

### Converter (Selected Date)

#### Calendar
- **Calendar**
  - Select the type of calendar system (e.g., Gregorian, Hijri, Umm Al-Qura).

#### Date Conversion
- **From Source Date Format**
  - Select the format of the input date.

- **Switch**
  - Reverse the conversion between the selected calendar systems.

- **To Target Date Format**
  - Select the target format for conversion.

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
  - Open a date picker to select and insert a specific Kurdish date based on Gregorian calendar.


### COM Library Integration
- The Kurdish Digital Calendar (KDC) provides a COM Library to facilitate the use of Kurdish date and number conversion in various Office applications through VBA. This guide explains how to use the KDC COM Library in VBA.
- The KDC COM Library provides a set of functions for converting dates between different calendar systems and formats and converting numbers to Kurdish text:

	- **ToKurdish** (Support both Kurdish Central and Kurdish Northern Dialects)
		- Insert the current Kurdish date into the document.
		- Syntax: `ToKurdish(formatChoice, dialect, isAddSuffix)` Default Dialect is Kurdish Central.
			- Parameters
				- formatChoice: The format of the output date.
				- dialect: The dialect of Kurdish to use ("ckb" for Central or "ku" for Northern).
				- isAddSuffix: A boolean value indicating whether to add suffixes to the output date.
			- Example: `ToKurdish(1, "Kurdish (Central)", true)` or `ToKurdish(1, "Kurdish (Northern)", true)`
			- Result: 
				- Kurdish Central Dialect : "دووشەممە، 11 بەفرانبار، 2723ی كوردی"
				- Kurdish Northern Dialect : "Duşem, 11 Berfanbar, 2723 Kurdî"
			- Note: The function supports dates from 21/03/0001 to 31/12/9999.
			
	- **ConvertNumberToKurdishCentralText**
		- Converts a number to Kurdish Central text.
		- Syntax: `ConvertNumberToKurdishCentralText(Number)`
			- Parameters
				- Number: The number to convert to Kurdish text.
			- Example: `ConvertNumberToKurdishCentralText(123456789)`
			- Result: "سەد و بیست و سێ ملیۆن و چوار سەد و پەنجا و شەش هەزار و حەوت سه‌د و هه‌شتاو نۆ"
			- Note: The function supports numbers up to 999,999,999,999,999,999.
			
	- **ConvertNumberToKurdishNorthernText**
		- Converts a number to Kurdish Northern text.
		- Syntax: `ConvertNumberToKurdishNorthernText(Number)`
			- Parameters
				- Number: The number to convert to Kurdish text.
			- Example: `ConvertNumberToKurdishNorthernText(123456789)`
			- Result: "sed û bîst û sê milyon û çar sed û pêncî û şeş hezar û heft sed û heştê û neh"
			- Note: The function supports numbers up to 999,999,999,999,999,999.
				
	- **ConvertDateBasedOnUserSelection** (Support both Kurdish Central and Kurdish Northern Dialects)
		- Converts a date between different calendar systems and formats based on user selection.
		- Syntax: `ConvertDateBasedOnUserSelection(Date, fromCalendar, toCalendar, fromFormat, toFormat, targetDialect, isAddSuffix)` Default Dialect is Kurdish Central.
		- Parameters
			- Date: The date to convert.
			- fromCalendar: The calendar system of the input date.
			- toCalendar: The calendar system of the output date.
			- fromFormat: The format of the input date.
			- toFormat: The format of the output date.
			- targetDialect: The dialect of Kurdish to use ("ckb" for Central or "ku" for Northern).
			- isAddSuffix: A boolean value indicating whether to add suffixes to the output date.
		- Example: `ConvertDateBasedOnUserSelection("01/01/2024", "Gregorian", "Kurdish", "dd/MM/yyyy", "dddd, dd MMMM, yyyy", "Kurdish (Central)", true)` or `KDC.ConvertDate("01/01/2024", "Gregorian", "Kurdish", "dd/MM/yyyy", "dddd, dd MMMM, yyyy", "Kurdish (Northern)", true)`
		- Result: 
			- Kurdish Central Dialect : "دووشەممە، 11 بەفرانبار، 2723ی كوردی"
			- Kurdish Northern Dialect : "Duşem, 11 Berfanbar, 2723 Kurdî"
		- Note: The function supports dates from 0002-01-01 to 9999-12-31.

#### Prerequisites
- Ensure you have the KDC COM Library installed and registered on your system.
- Add a reference to the KDC COM Library in your VBA editor.

#### Adding Reference to KDC COM Library
1. Open your VBA editor in Excel (or any other Office application).
2. Go to `Tools` > `References`.
3. Check the box for `Kurdish Digital Calendar Library`.

#### KDC COM Library Functions in VBA Code Example
```
' Declare a reference to the .NET class
Dim kdcService As Object

' Insert Now Kurdish Date
Function ToKurdish(formatChoice As Integer, dialect As String, isAddSuffix As Boolean) As String
    On Error GoTo ErrorHandler
    Set kdcService = CreateObject("KDCLibrary.KDCServiceImplementation")
    ToKurdish = kdcService.ToKurdish(formatChoice, dialect, isAddSuffix)
    Exit Function

ErrorHandler:
    ToKurdish = "Error: " & Err.Description
End Function

' Convert Date Based On User Selection
Function ConvertDateBasedOnUserSelection(selectedText As String, isReverse As Boolean, targetDialect As String, fromFormat As String, toFormat As String, targetCalendar As String, isAddSuffix As Boolean) As String
    On Error GoTo ErrorHandler
    Set kdcService = CreateObject("KDCLibrary.KDCServiceImplementation")
    ConvertDateBasedOnUserSelection = kdcService.ConvertDateBasedOnUserSelection(selectedText, isReverse, targetDialect, fromFormat, toFormat, targetCalendar, isAddSuffix)
    Exit Function

ErrorHandler:
    ConvertDateBasedOnUserSelection = "Error: " & Err.Description
End Function

' Convert Number To Kurdish Central Text
Function ConvertNumberToKurdishCentralText(number As Long) As String
    On Error GoTo ErrorHandler
    Set kdcService = CreateObject("KDCLibrary.KDCServiceImplementation")
    ConvertNumberToKurdishCentralText = kdcService.ConvertNumberToKurdishCentralText(number)
    Exit Function

ErrorHandler:
    ConvertNumberToKurdishCentralText = "Error: " & Err.Description
End Function

' Convert Number To Kurdish Northern Text
Function ConvertNumberToKurdishNorthernText(number As Long) As String
    On Error GoTo ErrorHandler
    Set kdcService = CreateObject("KDCLibrary.KDCServiceImplementation")
    ConvertNumberToKurdishNorthernText = kdcService.ConvertNumberToKurdishNorthernText(number)
    Exit Function

ErrorHandler:
    ConvertNumberToKurdishNorthernText = "Error: " & Err.Description
End Function

' Test the functions
Sub ExampleUsage()
    Dim number As Long
    number = 12345
    MsgBox "Kurdish Text (Central): " & ConvertNumberToKurdishCentralText(number)
    MsgBox "Kurdish Text (Northern): " & ConvertNumberToKurdishNorthernText(number)
    
    Dim kurdishDate As String
    kurdishDate = ToKurdish(1, "Kurdish (Central)", True)
    MsgBox "Kurdish Date: " & kurdishDate
    
    Dim convertedDate As String
    convertedDate = ConvertDateBasedOnUserSelection("01/01/2024", False, "Kurdish (Central)", "dd/MM/yyyy", "dddd, dd MMMM, yyyy", "Gregorian", True)
    MsgBox "Converted Date: " & convertedDate
End Sub
```


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
			    - targetDialect: The dialect of Kurdish to use ("ckb" for Central or "ku" for Northern).
			    - fromFormat: The format of the input date.
			    - toFormat: The format of the output date.
			    - targetCalendar: The calendar system to use for conversion.
			    - isAddSuffix (optional): A boolean value indicating whether to add suffixes to the output date.
            - Example: `ConvertDateToKurdish("01/01/2024", "Kurdish (Central)", "dd/MM/yyyy", "dddd, dd MMMM, yyyy", "Gregorian", true)` or `ConvertDateToKurdish(01/01/2024", "Kurdish (Northern)", "dd/MM/yyyy", "dddd, dd MMMM, yyyy", "Gregorian", true)` Default Dialect is Kurdish Central.
            - Result: 
				- Kurdish Central Dialect : "دووشەممە، 11 بەفرانبار، 2723ی كوردی"
				- Kurdish Northern Dialect : "Duşem, 11 Berfanbar, 2723 Kurdî"
            - Note: The function supports dates from 21/03/0001 to 31/12/9999.

## Installation

### Prerequisites

- Microsoft Office (Word, Excel, PowerPoint, Outlook, Access, Visio, Project)
- .NET Framework 4.7.2 or higher
- Visual Studio Tools for Office (VSTO) Runtime 2010 or higher

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

This project is licensed under the Custom License for Kurdish Digital Calendar (KDC).

You are free to use this software for personal, educational, or internal business purposes only. Redistribution, publication, and commercial use are strictly prohibited without express written permission from the author. For more details, please refer to the [LICENSE](https://github.com/6ebeng/kurdish-digital-calendar/blob/master/LICENSE) file.

For permissions beyond the scope of this license, please contact [rekbin.devs@gmail.com](mailto:rekbin.devs@gmail.com).


## Credits

- Developed and maintained by Tishko Rasoul ([Rekbin Devs](https://github.com/Rekbin-Devs)). 
- The kurdish calendar algorithm is based on the following sources:
  - [Kurdish Academy Calendar (2009)](https://www.kurdipedia.org/default.aspx?q=2013021510204775788&lng=5)
  - [Kurdish Calendar from Wikipedia](https://ckb.wikipedia.org/wiki/%DA%95%DB%86%DA%98%DA%98%D9%85%DB%8E%D8%B1%DB%8C_%DA%A9%D9%88%D8%B1%D8%AF%DB%8C#%D8%A8%DB%95%D8%B1%D8%A7%D9%88%D8%B1%D8%AF%DB%8C_%D8%B3%D8%A7%DA%B5%DB%8C_%DA%A9%D9%88%D8%B1%D8%AF%DB%8C_%D9%84%DB%95%DA%AF%DB%95%DA%B5_%D8%B3%D8%A7%DA%B5%DB%8C_%D8%B2%D8%A7%DB%8C%DB%8C%D9%86%DB%8C%D8%AF%D8%A7)
  - [Hjri Calendar from Wikipedia](https://ckb.wikipedia.org/wiki/%DA%95%DB%86%DA%98%DA%98%D9%85%DB%8E%D8%B1%DB%8C_%DA%A9%DB%86%DA%86%DB%8C)
  - [Salname](https://ku.wikipedia.org/wiki/Salname)
  - Thanks to all contributors and supporters.

## Support

For issues or questions, open an issue on the [Kurish Digital Calendar](https://github.com/6ebeng/kurdish-digital-calendar/issues) or contact us at [rekbin.devs@gmail.com](mailto:rekbin.devs@gmail.com).
