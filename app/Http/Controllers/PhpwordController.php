<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Element\Section;
use PhpOffice\PhpWord\IOFactory;

class PhpwordController extends Controller
{
    //
    public function createWord()
    {
        // Create a new PHPWord object
        $phpWord = new PhpWord();
        // Add a section to the document
        $section = $phpWord->addSection();
        // Add content to the section
        $section->addText('Hello, this is the content of the Word document.');

        // Save the document to a temporary file
        $filename = public_path('file/output_document.docx');
        $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save($filename);

        // Return the file as a download response
        return response()->download($filename, 'output_document.docx');
    }

    public function modifyDocument()
    {
        $filePath = public_path("file/private.docx");
        // Create a new PHPWord object
        $phpWord = new \PhpOffice\PhpWord\TemplateProcessor($filePath);

        $phpWord->setValues([
            'name' => "Rhalim calnetic"
        ]);

        $section = $phpWord->getSection(0);

        // Add header and footer
        $header = $section->addHeader();
        $footer = $section->addFooter();

        // Set header content with an image
        $imagePath = public_path('images/logo.png'); // Replace with the actual path to your image
        $header->addImage($imagePath, array('width' => 250, 'height' => 60));

        // Set footer content with HTML
        $htmlFooterContent = '
                <h1 style="font-size: 16px; color: #00395d;">
                YOUR WEBSITE HERE | YOUR NUMBER HERE | YOUR EMAIL HERE
                </h1>
                
                <p style="text-align: center; font-size: 10px;">This document and its content are protected by Canadian, U.S. and International copyright laws. Reproduction and distribution <br/> of this document and its content without the written permission of <strong>Your Company Name</strong> Here is strictly prohibited.
                </p>';
        \PhpOffice\PhpWord\Shared\Html::addHtml($footer, $htmlFooterContent, false, false);

        $phpWord->saveAs(public_path("file/hasilEdit.docx"));
    }

    public function generateWord()
    {
        $filePath = public_path("file/document.docx");
        // $filePath = "file/output_document.docx";
        echo $filePath . "<br>";

        $tempDir = sys_get_temp_dir();
        // var_dump($tempDir);
        \PhpOffice\PhpWord\Settings::setTempDir($tempDir);

        if (file_exists($filePath)) {

            try {
                $phpWord = IOFactory::load($filePath);

                var_dump($phpWord); die;

                // Check if there are sections in the loaded document
                $sections = $phpWord->getSections();
                if (empty($sections)) {
                    throw new Exception('The loaded document has no sections.');
                }

                // // Add a header to the first section
                $section = $sections[0];
                $header = $section->addHeader();
                $table = $header->addTable();
                $table->addRow();
                $cell = $table->addCell();
                $cell->addText('Your Header Text Here');

                // // Save the modified document
                $outputFile = public_path('file/file_modified.docx');
                $phpWord->save($outputFile);

                // echo "Header added successfully to the DOCX file.\n";
            } catch (\PhpOffice\PhpWord\Exception\Exception $e) {
                echo "Error loading the DOCX file: " . $e->getMessage() . "\n";
            } catch (Exception $e) {
                echo "Error: " . $e->getMessage() . "\n";
            }

            // try {
            //     // Load the Word document
            //     // $phpWord = \PhpOffice\PhpWord\IOFactory::load($filePath);

            //     \PhpOffice\PhpWord\Settings::setCompatibility(true);
            //     $phpWord = IOFactory::createReader('Word2007')
            //         ->load($filePath);

            //     if ($phpWord instanceof PhpOffice\PhpWord\PhpWord) {
            //         // The file was loaded successfully
            //         echo 'File loaded successfully!' . PHP_EOL;

            //         // Access document properties (optional)
            //         $properties = $phpWord->getDocInfo();
            //         echo 'Title: ' . $properties->getTitle() . PHP_EOL;
            //         echo 'Creator: ' . $properties->getCreator() . PHP_EOL;
            //         echo 'Last modified by: ' . $properties->getLastModifiedBy() . PHP_EOL;
            //         echo 'Created: ' . $properties->getCreated() . PHP_EOL;
            //         echo 'Modified: ' . $properties->getModified() . PHP_EOL;

            //         // Access document content (optional)
            //         $sections = $phpWord->getSections();
            //         foreach ($sections as $section) {
            //             $elements = $section->getElements();
            //             foreach ($elements as $element) {
            //                 // Process each element as needed
            //                 // For example, if $element is a \PhpOffice\PhpWord\Element\TextRun, you can access its text using $element->getText()
            //             }
            //         }
            //     } else {
            //         echo 'Error: Unable to load the file.' . PHP_EOL;
            //     }

            //     die;

            //     // var_dump($phpWord); die;
            //     // $phpWord = \PhpOffice\PhpWord\IOFactory::load($filePath);

            //     $section = $phpWord->getSection(0);

            //     // Add header and footer
            //     // $header = $section->addHeader();
            //     $footer = $section->addFooter();

            //     // Set header content with an image
            //     // $imagePath = public_path('images/logo.png'); // Replace with the actual path to your image
            //     // $header->addImage($imagePath, array('width' => 250, 'height' => 60));

            //     // Set footer content with HTML
            //     $htmlFooterContent = '
            //     <h1 style="font-size: 16px; color: #00395d;">
            //     YOUR WEBSITE HERE | YOUR NUMBER HERE | YOUR EMAIL HERE
            //     </h1>

            //     <p style="text-align: center; font-size: 10px;">This document and its content are protected by Canadian, U.S. and International copyright laws. Reproduction and distribution <br/> of this document and its content without the written permission of <strong>Your Company Name</strong> Here is strictly prohibited.
            //     </p>';
            //     \PhpOffice\PhpWord\Shared\Html::addHtml($footer, $htmlFooterContent, false, false);

            //     // Save the Word document
            //     $filename = 'document-result.docx';
            //     $path = public_path("file/{$filename}");
            //     $phpWord->save($path);

            //     // $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
            //     // $objWriter->save($path);

            //     // Download the Word document
            //     return response()->download($path, $filename);

            //     // Check if $phpWord is an instance of PhpOffice\PhpWord\PhpWord

            // } catch (Exception $e) {
            //     echo 'Error: ' . $e->getMessage() . PHP_EOL;
            // }
        } else {
            return response()->json(['error' => 'File not found'], 404);
        }
    }
}
