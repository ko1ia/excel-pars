<!doctype html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Test</title>
</head>
<body>
<form enctype="multipart/form-data" method="post">
    <input type="file" name="file">
    <input type="submit">
</form>

<?php

function dump($data) {
    echo '<pre>';
    print_r($data);
    echo '</pre>';
}

function addDesc($i, $sheet) {
    $mass = [];
    $var = $i + 5;
    for($i; $i < $var; $i++) {
        array_push($mass, $sheet->getCell('F'.$i)->getValue());
    }
    array_shift($mass);
    return $mass;
}

function countRow($sheet, $excel) {
    $sheet_info = $excel->getActiveSheet()->getHighestRowAndColumn();
    $skill = [];
    for($i = 0; $i < $sheet_info['row']; $i++) {
        if($sheet->getCell('A'.$i)->getValue() != NULL &&
            strlen($sheet->getCell('A'.$i)->getValue())< 3) {
            array_push($skill, [$sheet->getCell('A'.$i)->getValue(),$sheet->getCell('A'.$i)->getRow()]);
        }
    }
    return $skill;
}

$result = [];

if (isset($_POST)) {
    require_once 'php/PHPExcel.php';

    $random = substr(md5(mt_rand()), 0, 7);
    $uploadfile = 'files/' . $random . $_FILES['file']['name'];
    if (move_uploaded_file($_FILES['file']['tmp_name'], $uploadfile)) {
        $excel = PHPExcel_IOFactory::load($uploadfile);
        $excel->setActiveSheetIndex(0);
        $sheet = $excel->getActiveSheet();
        $skill = countRow($sheet, $excel);

        $k = 0;
        $mass = [];
        for($i = $skill[$k][1]; $i < $skill[count($skill)-1][1]; $i++){
            if($i < $skill[$k+1][1]) {
                $itter = $i + 1;
                if($sheet->getCell('C'.$itter)->getValue() != NULL && strlen($sheet->getCell('C'.$itter)->getValue()) == 1) {
                    if($sheet->getCell('B'.$i)->getValue() != NULL) {
                        array_push($mass,
                        [
                            'sub' =>        $sheet->getCell('B'.$i)->getValue(),
                            'type' =>       $sheet->getCell('C'.$itter)->getValue(),
                            'aspect'=>      $sheet->getCell('D'.$itter)->getValue(),
                            'description'=> $sheet->getCell('C'.$itter)->getValue() == 'J' ? addDesc($itter, $sheet) :  $sheet->getCell('F'. $itter)->getValue(),
                            'max' =>        $sheet->getCell('I'.$itter)->getValue(),
                        ]);
                    } else {
                        array_push($mass,
                        [
                            'type' =>       $sheet->getCell('C'.$itter)->getValue(),
                            'aspect'=>      $sheet->getCell('D'.$itter)->getValue(),
                            'description'=> $sheet->getCell('C'.$itter)->getValue() == 'J' ? addDesc($itter, $sheet) :  $sheet->getCell('F'. $itter)->getValue(),
                            'max' =>        $sheet->getCell('I'.$itter)->getValue(),
                        ]);
                    }
                    
                    
                }

                if($sheet->getCell('C'.$itter)->getValue() == 'J') {
                    $i += 4;
                }

                if($i == $skill[$k+1][1]-1) {
                    array_push($result, [$skill[$k][0] => $mass]);
                    $mass = [];
                    $k++;
                }
            }
        }
    }
    dump($result);
}

?>
</body>
</html>