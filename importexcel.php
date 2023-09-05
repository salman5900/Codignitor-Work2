<!DOCTYPE html>
<html>
<head>
    <title>Excel Import</title>
</head>
<body>
    <h1>Excel Import</h1>

    <form method="post" enctype="multipart/form-data">
        <input type="file" name="excelFile" accept=".xlsx, .xls">
        <input type="submit" name="submit" value="Upload and Modify">
    </form>

    <?php
    require 'vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\IOFactory;
    use PhpOffice\PhpSpreadsheet\Spreadsheet;


    $highestRow = 0;
    $highestColumnIndex = 0;
    $spreadsheet = new Spreadsheet();

    if (isset($_POST['submit'])) {
        if (isset($_FILES['excelFile']) && $_FILES['excelFile']['error'] === UPLOAD_ERR_OK) {
            $uploadedFilePath = $_FILES['excelFile']['tmp_name'];

            try {
            
                $spreadsheet = IOFactory::load($uploadedFilePath);
                $worksheet = $spreadsheet->getActiveSheet();

              
                $highestRow = $worksheet->getHighestRow();
                $highestColumn = $worksheet->getHighestColumn();
                $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);

            
                echo '<h2>Excel Content</h2>';
                echo '<form method="post">';
                echo '<table border="1">';
                for ($row = 1; $row <= $highestRow; $row++) {
                    echo '<tr>';
                    for ($col = 1; $col <= $highestColumnIndex; $col++) {
                        $cellValue = $worksheet->getCellByColumnAndRow($col, $row)->getValue();
                        echo "<td>Row: $row, Column: $col - Value: $cellValue</td>";
                    }
                    echo '<td><select name="dropdown[]">';
                    echo '<option value="Option 1">Option 1</option>';
                    echo '<option value="Option 2">Option 2</option>';
                    echo '<option value="Option 3">Option 3</option>';
                    echo '</select></td>';
                    echo '</tr>';
                }
                echo '</table>';
                echo '<input type="submit" name="saveButton" value="Save Excel File">';
                echo '</form>';
            } catch (Exception $e) {
                echo 'Error: ' . $e->getMessage();
            }
        } else {
            echo 'Please choose a valid Excel file to upload.';
        }
        echo "Loaded Excel file from: $uploadedFilePath<br>";
    }

 // ...

if (isset($_POST['saveButton'])) { 

    $dropdownValues = $_POST['dropdown'];
    for ($row = 1; $row <= $highestRow; $row++) {
        $selectedValue = $dropdownValues[$row - 1];
        $worksheet->setCellValueByColumnAndRow($highestColumnIndex + 1, $row, $selectedValue);
    }

    
    $newExcelFilePath = __DIR__ . '\modified_excel.xlsx';
    try {
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($newExcelFilePath);
        for ($row = 1; $row <= $highestRow; $row++) {
            echo "Row: $row - Dropdown Value: {$dropdownValues[$row - 1]}<br>";}

        echo "Modified Excel file saved: <a href='$newExcelFilePath' target='_blank'>$newExcelFilePath</a>";
        echo "Attempting to save Excel file to: $newExcelFilePath<br>";
        
    } catch (Exception $e) {
        echo 'Error while saving the Excel file: ' . $e->getMessage();
    }
}

    ?>
</body>
</html>
