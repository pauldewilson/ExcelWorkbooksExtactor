# ExcelWorkbooksExtractor

#### Compare specific cell values of a specific sheet of different workbooks

If you have many Excel workbooks and need to:
<ol>
<li> Compare the same cell (or cells) values across various workbooks (like cell A1 value across 1,000 workbooks)
<li> Extract a specific cell's value to compare all the files' values in one place
</ol>

Then this script will perform this task.

# Example Input

The included sample_files folder contains 3x xlsx files which contain the below.

#### sample_one.xlsx
<table>
<th></th>
<th>A</th>
<th>B</th>
<tr>
<td>
1
</td>
<td>
11111_a1
</td>
<td>

</td>
</tr>
<tr>
<td>
2
</td>
<td>

</td>
<td>
11111_b2
</td>
</tr>
</table>

#### sample_two.xlsx
<table>
<th></th>
<th>A</th>
<th>B</th>
<tr>
<td>
1
</td>
<td>
22222_a1
</td>
<td>

</td>
</tr>
<tr>
<td>
2
</td>
<td>

</td>
<td>
22222_b2
</td>
</tr>
</table>

#### sample_three.xlsx
<table>
<th></th>
<th>A</th>
<th>B</th>
<tr>
<td>
1
</td>
<td>
33333_a1
</td>
<td>

</td>
</tr>
<tr>
<td>
2
</td>
<td>

</td>
<td>
33333_b2
</td>
</tr>
</table>

# Example Output

#### output.csv
Note: Column headers are the File Name followed by the target cells
<table>
<th>
File Name
</th>
<th>
A1
</th>
<th>
B2
</th>
<tr>
<td>
sample_one.xlsx
</td>
<td>
11111_a1
</td>
<td>
11111_b2
</td>
</tr>
<tr>
<td>
sample_two.xlsx
</td>
<td>
22222_a1
</td>
<td>
22222_b2
</td>
</tr>
<tr>
<td>
sample_three.xlsx
</td>
<td>
33333_a1
</td>
<td>
33333_b2
</td>
</tr>
</table>

# How to

<ol>
<li>Open <i>target_cells.txt</i>
<li>List all the target cells, each on a new line.
<li>Save and close the file
<li>Run app.py
<li>Input the target directory path where your Excel workbooks are located
<li>Type the target sheet name (case sensitive)
<li>Done. `output.csv` will be output into the same directory you provided with the workbooks.
<li>Any errors will be printed to the terminal for you to examine
</ol>

# Compatible filetypes
.xlsx
.xlsm
.xltx
.xltm
<br>
If you want .csv for the time being please save as .xlsx.
<br>
Happy to implement this feature if demand exists.