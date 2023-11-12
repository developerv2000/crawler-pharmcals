<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Str;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class MainController extends Controller
{
    public function index()
    {
        $year = 2023;
        $fromPage = 1;
        $toPage = 119;

        for ($i = $fromPage; $i <= $toPage; $i++) {
            $response = Http::withBasicAuth('Ortos', 'Ortos2023')->get('http://pharm.cals.am/pharm/report/get_data.php', [
                'prand' => (float)rand() / (float)getrandmax(),
                'pbtn' => 'search',
                'pdate1' => '01-01-' . $year,
                'pdate2' => '31-12-' . $year,
                'pname' => '',
                'pgeneric' => '',
                'pdosform' => '',
                'pcountry' => '',
                'pmanuf' => '',
                'ptype' => '1',
                'ppage' => $i,
                'psid' => '611212656',
            ]);

            if ($response->successful()) {
                $result = json_decode($response->body());

                // preapare excel reader/writer
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load(public_path('results.xlsx'));
                $sheet = $spreadsheet->getActiveSheet();

                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
                $reader->setReadDataOnly(true);

                $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
                $writer->setPreCalculateFormulas(false);

                for ($j = 0; $j < count($result->items); $j++) {
                    $drug = $this->explodeeDrug($result->items[$j]->caption);
                    $count = $result->items[$j]->count;

                    $highestRow = $sheet->getHighestRow() + 1;
                    // add data to excel file
                    $sheet->setCellValue('A' . $highestRow, $drug[0]); // name
                    $sheet->setCellValue('B' . $highestRow, $drug[3]); // composition
                    $sheet->setCellValue('C' . $highestRow, $drug[1]); // dosage
                    $sheet->setCellValue('D' . $highestRow, $drug[2]); // address
                    $sheet->setCellValue('E' . $highestRow, $count); // count
                    $sheet->setCellValue('F' . $highestRow, $year); // year
                }

                // Save generated file
                $writer->save(public_path('results.xlsx'));
            } else {
                dd('Error while parsing page ' . ($i + 1) . ', year: ' . $year);
            }
        }
        return 'Success!';
    }

    private function explodeeDrug($string): array
    {
        $exploded = explode('<br>', $string);

        for($k = 0; $k < count($exploded); $k++) {
            $exploded[$k] = strip_tags($exploded[$k]);
            $exploded[$k] = Str::squish($exploded[$k]);
        }

        // generate text for composition column from name
        $name = $exploded[0];
        preg_match_all('#\(((?>[^()]+)|(?R))*\)#x', $name, $composition);
        $matches = [];

        // ignore unwanted items like '(+)', '(-)', '(1)'
        foreach ($composition[0] as $item) {
            if(strlen($item) > 3) {
                // remove () from string
                $matches[] = substr($item, 1, strlen($item) - 2);
            }
        }

        $exploded[3] = implode(' + ', $matches);

        return $exploded;
    }
}
