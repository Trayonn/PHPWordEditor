<?php
require 'vendor/autoload.php';

use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\IOFactory;
use Dompdf\Dompdf;
use Dompdf\Options;


$tipo = null;
$QTD_ITENS = $_POST['QTD_ITENS'];
$VALOR_CONTRATO = null;

if($QTD_ITENS <= 1000 ){
    $VALOR_CONTRATO = 1200.00;
    $tipo = 1;
}
elseif($QTD_ITENS > 1000 && $QTD_ITENS <= 5000){
    $VALOR_CONTRATO = 1800.00;
    $tipo = 2;
}
elseif($QTD_ITENS > 5000 && $QTD_ITENS <= 10000 ){
    $VALOR_CONTRATO = 3000.00;
    $tipo = 3;
}
elseif($QTD_ITENS > 10000 && $QTD_ITENS <= 20000 ){
    $VALOR_CONTRATO = 6000.00;
    $tipo = 4;
}
elseif($QTD_ITENS > 20000 ){
    $VALOR_CONTRATO = 8400.00;
    $tipo = 5;
}



$templateProcessor = new TemplateProcessor("Contrato_Modelo_TEMPLATE.docx");



$templateProcessor->setValue('CONTRATANTE', $_POST["CONTRATANTE"]);
$templateProcessor->setValue('CONTRATADA', $_POST["CONTRATADA"]);
$templateProcessor->setValue('NOME', $_POST["NOME"]);
$templateProcessor->setValue('EMAIL', $_POST["EMAIL"]);
$templateProcessor->setValue('TELEFONE', $_POST["TELEFONE"]);
$templateProcessor->setValue('QTD_ITENS', $QTD_ITENS);
$templateProcessor->setValue('TIPO_CONTRATO', $tipo);
$templateProcessor->setValue('VALOR_CONTRATO', $VALOR_CONTRATO);

$filePath = 'Contrato_Gerado.docx';
$templateProcessor->saveAs($filePath);

$formato = $_POST['formato'];

if ($formato === 'pdf') {
    $phpWord = IOFactory::load($filePath);
    $htmlConverter = IOFactory::createWriter($phpWord, 'HTML');
    ob_start();

    $htmlConverter->save('php://output');
    $htmlConteudo = ob_get_clean();
    $options = new Options();
    $options->set('defaultFont', 'Arial');
    $dompdf = new Dompdf($options);
    $dompdf->loadHtml($htmlConteudo);
    $dompdf->setPaper('A4', 'portrait');
    $dompdf->render();

    $filePathPDF = 'Contrato_Gerado.pdf';
    file_put_contents($filePathPDF, $dompdf->output());
    header('Content-Type: application/pdf');
    header('Content-Disposition: attachment; filename="Contrato_Gerado.pdf"');
    readfile($filePathPDF);

} else {
    header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    header('Content-Disposition: attachment; filename="Contrato_Gerado.docx"');
    header('Content-Length: ' . filesize($filePath));
    readfile($filePath);
}

?>
