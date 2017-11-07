<?php

namespace AppBundle\Controller;

use AppBundle\Entity\ExcelFile;
use AppBundle\Form\Type\ExcelFormatType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Csv as ReaderCsv;
use PhpOffice\PhpSpreadsheet\Reader\Ods as ReaderOds;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as ReaderXlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Ods;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Sensio\Bundle\FrameworkExtraBundle\Configuration\Route;
use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\StreamedResponse;

class DefaultController extends Controller
{
    const FILENAME = 'Browser_characteristics';

    /**
     * @Route("/", name="homepage")
     */
    public function indexAction(Request $request) {
        return $this->render('AppBundle::index.html.twig');
    }

    /**
     * @Route("/export", name="export")
     */
    public function exportAction(Request $request)
    {
        $form = $this->createForm(ExcelFormatType::class);
        $form->handleRequest($request);
        if ($form->isSubmitted() && $form->isValid()) {
            $data = $form->getData();
            $format = $data['format'];
            $filename = self::FILENAME.'.'.$format;

            $spreadsheet = $this->createSpreadsheet();

            switch ($format) {
                case 'ods':
                    $writer = new Ods($spreadsheet);
                    break;
                case 'xlsx':
                    $writer = new Xlsx($spreadsheet);
                    break;
                case 'csv':
                    $writer = new Csv($spreadsheet);
                    break;
                default:
                    return $this->render('AppBundle::export.html.twig', [
                        'form' => $form->createView(),
                    ]);
            }

            $response = new StreamedResponse();
            $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            $response->headers->set('Content-Disposition', 'attachment;filename="'.$filename.'"');
            $response->setPrivate();
            $response->headers->addCacheControlDirective('no-cache', true);
            $response->headers->addCacheControlDirective('must-revalidate', true);
            $response->setCallback(function() use ($writer) {
                $writer->save('php://output');
            });

            return $response;
        }

        return $this->render('AppBundle::export.html.twig', [
            'form' => $form->createView(),
        ]);
    }

    protected function createSpreadsheet() {
        $spreadsheet = new Spreadsheet();
        // Get active sheet - also possible to retrieve a specific sheet
        $sheet = $spreadsheet->getActiveSheet();

        // Set cell name and merge cells
        $sheet->setCellValue('A1', 'Browser characteristics')->mergeCells('A1:D1');

        // Set column names
        $columnNames = [
            'Browser',
            'Developper',
            'Release date',
            'Written in',
        ];
        $columnLetter = 'A';
        foreach ($columnNames as $columnName) {
            // Allow to access AA column if needed and more
            $sheet->setCellValue($columnLetter++.'2', $columnName);
        }

        // Add data for each column
        $columnValues = [
            ['Google Chrome', 'Google Inc.', 'September 2, 2008', 'C++'],
            ['Firefox', 'Mozilla Foundation', 'September 23, 2002', 'C++, JavaScript, C, HTML, Rust'],
            ['Microsoft Edge', 'Microsoft', 'July 29, 2015', 'C++'],
            ['Safari', 'Apple', 'January 7, 2003', 'C++, Objective-C'],
            ['Opera', 'Opera Software', '1994', 'C++'],
            ['Maxthon', 'Maxthon International Ltd', 'July 23, 2007', 'C++'],
            ['Flock', 'Flock Inc.', '2005', 'C++, XML, XBL, JavaScript'],
        ];

        $i = 3; // Beginning row for active sheet
        foreach ($columnValues as $key => $columnValue) {
            $columnLetter = 'A';
            foreach($columnValue as $k => $v) {
                $sheet->setCellValue($columnLetter++.$i, $v);
            }
            $i++;
        }

        // Autosize each column and set style to column titles
        $columnLetter = 'A';
        foreach ($columnNames as $columnName) {
            // Center text
            $sheet->getStyle($columnLetter.'1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle($columnLetter.'2')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            // Text in bold
            $sheet->getStyle($columnLetter.'1')->getFont()->setBold(true);
            $sheet->getStyle($columnLetter.'2')->getFont()->setBold(true);
            // Autosize column
            $sheet->getColumnDimension($columnLetter)->setAutoSize(true);
            $columnLetter++;
        }

        return $spreadsheet;
    }

    /**
     * @Route("/import", name="import")
     */
    public function importAction(Request $request)
    {
        $filename = $this->get('kernel')->getRootDir().'/../export/'.self::FILENAME.'.xlsx';
        if (!file_exists($filename)) {
            exit('File does not exist.');
        }

        $spreadsheet = $this->readFile($filename);
        $data = $this->createDataFromSpreadsheet($spreadsheet);

        return $this->render('AppBundle::import.html.twig', [
            'data' => $data,
        ]);
    }

    protected function loadFile($filename) {
        return IOFactory::load($filename);
    }

    protected function readFile($filename) {
        $extension = pathinfo($filename, PATHINFO_EXTENSION);
        switch ($extension) {
            case 'ods':
                $reader = new ReaderOds();
                break;
            case 'xlsx':
                $reader = new ReaderXlsx();
                break;
            case 'csv':
                $reader = new ReaderCsv();
                break;
            default:
                throw new \Exception('Invalid extension');
        }
        $reader->setReadDataOnly(true);
        return $reader->load($filename);
    }

    protected function createDataFromSpreadsheet($spreadsheet) {
        $data = [];
        foreach ($spreadsheet->getWorksheetIterator() as $worksheet) {
            $worksheetTitle = $worksheet->getTitle();
            $data[$worksheetTitle] = [
                'columnNames' => [],
                'columnValues' => [],
            ];
            foreach ($worksheet->getRowIterator() as $row) {
                $rowIndex = $row->getRowIndex();
                if ($rowIndex > 2) {
                    $data[$worksheetTitle]['columnValues'][$rowIndex] = [];
                }
                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(false); // Loop all cells, even if it is not set
                foreach ($cellIterator as $cell) {
                    if ($rowIndex === 2) {
                        $data[$worksheetTitle]['columnNames'][] = $cell->getCalculatedValue();
                    }
                    if ($rowIndex > 2) {
                        $data[$worksheetTitle]['columnValues'][$rowIndex][] = $cell->getCalculatedValue();
                    }
                }
            }
        }

        return $data;
    }
}
