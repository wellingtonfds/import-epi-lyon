<?php

use App\epi;
use App\Imports\UsersImport;
use Illuminate\Support\Arr;
use Illuminate\Support\Facades\Artisan;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Route;
use Maatwebsite\Excel\Facades\Excel;
use Symfony\Component\Process\Process;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Artisan::command('import-file', function () {
    $this->comment('connnet vpn');
    $user = env('VPN_USER');
    $pass = env('VPN_PASS');
    $domain = env('VPN_DOMAIN');
    $url = env('VPN_URL');
    $process =  Process::fromShellCommandline("echo Y | netExtender -u $user  -p $pass -d $domain  --auto-reconnect $url &>/dev/null &");
    $process->start();
    $this->comment('wait connection');
    sleep(10);
    $this->comment('start processing');
    DB::table('epis')->truncate();
    $this->info('start import');
    $inputFileName = storage_path('app/public/EPIContrato.xlsx');
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $spreadsheet = $reader->load($inputFileName);
    $all = $spreadsheet->getSheetNames();

    $spreadsheetCount = $this->output->createProgressBar(count($all));
    $spreadsheetCount->start();
    $listEpi = [];
    $setPosition = '';

    foreach ($all as $name) {

        $epiList = [
            'epis' => [],
            'uniforme' => [],
        ];
        $worksheet = $spreadsheet->getSheetByName($name);

        foreach ($worksheet->getRowIterator() as $row) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(FALSE); // This loops through all cells,
            foreach ($cellIterator as $cell) {
                $indexCell = @$cell->getCoordinate()[0];
                if ($indexCell != 'B'  || empty($cell->getValue())) {
                    continue;
                }
                if ($cell->getValue() == "EPI's") {
                    $setPosition = 'epis';
                } elseif ($cell->getValue() == "Uniforme") {
                    $setPosition = 'uniforme';
                } else {
                    array_push($epiList[$setPosition], $cell->getValue());
                }
            }
        }
        $listEpi[] = [
            'cc' => $name,
            'meta' => $epiList
        ];
        $spreadsheetCount->advance();
    }
    $spreadsheetCount->finish();
    $setPosition = '';
    $epiList = [
        'epis' => [],
        'uniforme' => [],
    ];

    foreach ($listEpi as $epi) {
        epi::create([
            'cc' => $epi['cc'],
            'meta' => json_encode($epi['meta'])
        ]);
    }
    $process->stop(9, SIGINT);
    $this->comment('finish import');
});

Route::get('/', function () {

    $inputFileName = storage_path('app/public/EPIContrato.xlsx');
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $spreadsheet = $reader->load($inputFileName);
    $all = $spreadsheet->getSheetNames();

    $listEpi = [];
    $setPosition = '';

    foreach ($all as $name) {
        $epiList = [
            'epis' => [],
            'uniforme' => [],
        ];
        $worksheet = $spreadsheet->getSheetByName($name);

        foreach ($worksheet->getRowIterator() as $row) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(FALSE); // This loops through all cells,
            foreach ($cellIterator as $cell) {
                $indexCell = @$cell->getCoordinate()[0];
                if ($indexCell != 'B'  || empty($cell->getValue())) {
                    continue;
                }
                if ($cell->getValue() == "EPI's") {
                    $setPosition = 'epis';
                } elseif ($cell->getValue() == "Uniforme") {
                    $setPosition = 'uniforme';
                } else {
                    array_push($epiList[$setPosition], $cell->getValue());
                }
            }
        }
        $listEpi[] = [
            'cc' => $name,
            'meta' => $epiList
        ];
    }
    $setPosition = '';
    $epiList = [
        'epis' => [],
        'uniforme' => [],
    ];

    foreach ($listEpi as $epi) {
        epi::create([
            'cc' => $epi['cc'],
            'meta' => json_encode($epi['meta'])
        ]);
    }
    exit();
});
