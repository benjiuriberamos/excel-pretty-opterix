<?php
require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

// Read the file json
$json = file_get_contents('src/contracts.json');
$data = json_decode($json, true);

// Create a new Spreadsheet
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Define headers
$headers = array_keys($data[0]);
$sheet->fromArray($headers, NULL, 'A1');

// Add content
$row = 2; 
foreach ($data as $item) {
    $sheet->fromArray(array_values($item), NULL, 'A' . $row);
    $row++;
}

// Auto size each of columns
$columnLetter = 'A';
for ($i = 0; $i < count($headers); $i++) {
    $sheet->getColumnDimension($columnLetter)->setAutoSize(true);
    $columnLetter++;
}

// Add bold styles for headers
$sheet->getStyle('A1:' . $columnLetter . '1')->getFont()->setBold(true);

// Add align right styles for two first columns
$sheet->getStyle('A2:A' . $row)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);  // Primera columna
$sheet->getStyle('B2:B' . $row)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);  // Segunda columna

// Setting a background color
$color = true;
for ($i = 2; $i <= $row; $i++) {
    if ($color) {
        $sheet->getStyle('A' . $i . ':' . $columnLetter . $i)
            ->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('e2efd9'); // Background color e2efd9
    } else {
        $sheet->getStyle('A' . $i . ':' . $columnLetter . $i)
            ->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('b4c6e7'); // Background color b4c6e7
    }
    $color = !$color;
}

// Write the file contracts.xlsx
$writer = new Xlsx($spreadsheet);
$writer->save('contracts.xlsx');

echo "Excel exportado correctamente.";
