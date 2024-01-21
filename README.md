The prerequisites for using the provided PowerShell script are as follows:

1. **PowerShell Version:**
   Ensure that you have PowerShell installed on your system. The script should work with PowerShell 5.1 or later.

2. **ImportExcel Module:**
   The script uses the `ImportExcel` module to handle Excel files. If you don't have this module installed, the script attempts to install it automatically. However, you need to have the necessary permissions to install modules.

3. **OfficeOpenXml Assembly:**
   The script uses the `OfficeOpenXml` assembly for working with Excel files. If the assembly is not available, you may encounter errors. In the provided script, an attempt is made to load the assembly. If you face issues related to this assembly, you may need to manually install it or resolve any errors that prevent its loading.

4. **System.Windows.Forms Assembly:**
   The script uses the `System.Windows.Forms` assembly for GUI operations. This assembly is typically available on Windows systems. If you're running the script on a system without this assembly, you might need to install or enable the required .NET components.

5. **Internet Access (Optional):**
   If the script attempts to install the `ImportExcel` module from the PowerShell Gallery, it requires internet access. If internet access is restricted, you may need to manually install the module or adjust your network settings.

6. **Permissions:**
   Ensure that you have the necessary permissions to run PowerShell scripts on your system. If execution policies are restricted, you may need to adjust them using the `Set-ExecutionPolicy` cmdlet.

7. **Supported File Formats:**
   The script is designed to work with text files (`*.txt`), Excel files (`*.xlsx`), and CSV files (`*.csv`). Ensure that your files are in one of these formats.

8. **Excel File Structure (if applicable):**
   If you're working with Excel files, ensure that the file has a reasonable structure. The script assumes that each sheet in the Excel workbook contains data without column headers. Adjustments may be needed if your Excel files have a different structure.

By addressing these prerequisites, you should be able to use the script successfully. If you encounter any issues or have specific requirements, feel free to ask for assistance!




##############################
Usage
This PowerShell script enables users to extract and format data from various file formats, including text files, Excel spreadsheets, and CSV files.
The script prompts users to select a file, supports password-protected files, and extracts IPs or URLs. The extracted data is then saved in a well-organized structure within a folder named "FG_Formatted_Output," including double-quoted IPs or URLs and Fortigate-formatted data. The script simplifies the process of preparing data for network configurations or security policies.
