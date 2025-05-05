#  Montgomery Fleet Equipment Inventory Data Cleaning

This project involved cleaning and preparing the `Montgomery_Fleet_Equipment_Inventory_FA_PART_1_START.CSV` dataset using **Excel for the Web**. The cleaned version is saved as `Montgomery_Fleet_Equipment_Inventory_FA_PART_1_END.XLSX`.

##  Cleaning Tasks Performed

1. **File Conversion**
   - Converted the original CSV file to XLSX format after switching to *Editing* mode.
   - Used the "Convert" prompt in Excel to enable full Excel features.

2. **Column Formatting**
   - Resized all columns to ensure data is clearly visible in each cell.

3. **Empty Row Removal**
   - Used filter options to identify and remove all empty rows from the dataset.

4. **Duplicate Record Removal**
   - Identified and removed duplicate entries using Excel’s *Remove Duplicates* feature.

5. **Spelling Correction**
   - Manually reviewed and corrected spelling errors in cell values.

6. **Whitespace Cleanup**
   - Used *Find and Replace* to remove all instances of double-spaces in the dataset.

7. **Department Name Fix**
   - Department names were incorrectly split over two columns.
   - Merged the values into a single column using *Flash Fill*.
   - Verified correctness using a provided reference list.
   - Removed the now-unnecessary extra column.

8. **Final Export**
   - Used *Save As* → *Download a Copy* to export the cleaned workbook as:
     - `Montgomery_Fleet_Equipment_Inventory_FA_PART_1_END.XLSX`

---

✅ The cleaned dataset is now ready for analysis or further transformation.

# Montgomery Fleet Equipment Inventory – Part 2 Data Analysis

This phase involved formatting, analyzing, and visualizing the `Montgomery_Fleet_Equipment_Inventory_FA_PART_2_START.XLSX` file using **Excel for the Web**, with a focus on **Pivot Tables** and summary statistics. The cleaned and analyzed file was saved as `Montgomery_Fleet_Equipment_Inventory_FA_PART_2_END.XLSX`.

##  Tasks Performed

1. **Format as Table**
   - Used *Format as Table* to organize the dataset and enable structured references for further analysis.

2. **Summary Statistics for Column 'C' (Equipment Count)**
   - Applied AutoSum to compute:
     - **SUM:** Total of all values in Column C
     - **AVERAGE:** Mean value
     - **MIN:** Minimum value
     - **MAX:** Maximum value
     - **COUNT:** Number of entries

3. **Pivot Table 1 – Equipment Count by Department**
   - Created a Pivot Table with:
     - **Rows:** Department
     - **Values:** Sum of Equipment Count
   - Sorted the table in **descending order** by total Equipment Count.

4. **Pivot Table 2 – Equipment Count by Department and Class**
   - Duplicated Pivot Table 1.
   - Added **Equipment Class** *below* Department in the Rows section.
   - Collapsed all fields except for **Transportation** to show only its equipment breakdown.

5. **Pivot Table 3 – Equipment Count by Class and Department**
   - Duplicated Pivot Table 1 again.
   - Added **Equipment Class** *above* Department in the Rows section.
   - Collapsed all fields except for **CUV** to analyze its departmental breakdown.

6. **Final Export**
   - Saved and downloaded the completed workbook as:
     - `Montgomery_Fleet_Equipment_Inventory_FA_PART_2_END.XLSX`

---

✅ This completes the data analysis and pivot table exploration using Excel for the Web.

