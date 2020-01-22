<?php
setlocale(LC_NUMERIC, "en_US.UTF-8");

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

include 'vendor/autoload.php';

if (count($argv) == 1) {
    die('Usage: php generate.php exam.json' . "\n");
}

if (!$exam = @json_decode(@file_get_contents($argv[1]), true)) {
    die('Unable to parse json' . "\n");
}

$spreadsheet = new Spreadsheet;
$sheet = $spreadsheet->getActiveSheet();

$sheet->setCellValue('A2', 'Name');
$col = 2;
$totalWeight = 0;

foreach ($exam['questions'] as $entry) {
    list($name, $weight) = $entry;
    $sheet->getCellByColumnAndRow($col, 1)->setValue("$name");
    $sheet->getCellByColumnAndRow($col, 2)->setValue($weight);
    $totalWeight += $weight;
    $col++;
}

echo "Weights sum: $totalWeight\n";

$terms = [];
for ($i=2; $i<$col; $i++) {
    $cell = $sheet->getCellByColumnAndRow($i, 2);
    $term = '$'.$cell->getColumn().'$'.$cell->getRow();

    $grade = $sheet->getCellByColumnAndRow($i, 3);
    $term .= '*'.$grade->getCoordinate();
    $terms[] = $term;
}
$sheet->getCellByColumnAndRow($col, 3)->setValue('='.implode('+', $terms));
$total = 0;

$writer = new Xlsx($spreadsheet);
$writer->save('correction.xlsx');
