<?php

define('DS', DIRECTORY_SEPARATOR);

require_once __DIR__.DS.'vendor'.DS.'autoload.php';

if (strtolower($_SERVER['REQUEST_METHOD']) === 'post') {
    $file = $_FILES['file'] ?: null;
    $type = ($file['type'] == 'application/vnd.ms-excel') ? 'Excel5' : 'Excel2007';
    if ($type) {
        $version = $_POST['version'] ?: '3.0';
        $filename = $_POST['filename'] ?: 'SampleVCF.vcf';
        $excelReader = \PHPExcel_IOFactory::createReader($type);
        $excelReader->setReadDataOnly();
        $excelObj = $excelReader->load($file['tmp_name']);
        $excelObj->setActiveSheetIndex(0);
        $objWorksheet = $excelObj->getActiveSheet();
        $highestRow = $objWorksheet->getHighestRow();
        $highestColumn = $objWorksheet->getHighestColumn();
        $text = "";
        for ($i = 2; $i <= $highestRow; ++$i) {
            $row = $objWorksheet->rangeToArray('A'.$i.':E'.$i, null, true, true, false)[0];
            $firstName = $row[0];
            $lastName = $row[1];
            $mobileNumber = $row[2];
            $alternateNumber = $row[3];
            $email = $row[4];
            if ($firstName && $lastName && $mobileNumber) {
                $text .= sprintf("BEGIN:VCARD
VERSION:%s
N:%s;%s;;;
FN:%s
TEL;TYPE=CELL;TYPE=PREF:%s
", $version, $firstName, $lastName, trim($firstName).' '.trim($lastName), $mobileNumber, $alternateNumber, $email);
                if ($alternateNumber) {
                    $text .= sprintf("TEL;TYPE=WORK,VOICE:%s
", $alternateNumber);
                }
                if ($email) {
                    $text .= sprintf("EMAIL:%s
", $email);
                }
                $text .= "END:VCARD
";
            }
        }

        header('Content-Type: text/x-vcard');
        header("Content-Disposition: inline; filename=$filename");
        echo $text;
    }
} else {
?>
    <html>
        <body>
            <form action method="post" enctype="multipart/form-data">
                <input type="file" name="file">
                <input type="text" name="version" value="3.0">
                <input type="filename" name="filename" value="SampleVCF.vcf">
                <button type="submit">Convert to VCF</button>
            </form>
        </body>
    </html>
<?php
}

