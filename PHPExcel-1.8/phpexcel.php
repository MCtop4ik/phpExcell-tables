<?php

setlocale(LC_ALL, 'ru_RU.UTF-8', 'Russian_Russia.65001');

 $arr = array(
     "eyJCIiwiMCIsIjAiLCI1NTYiLCIwIiwiMCIsIiAgICAgICAgMTdEICAifQ == " => array(
         "name" => 'Коммерческий отдел',
         "groups" => array(
             "eyJCIiwiMCIsIjAiLCI5NjMzIiwiMCIsIjAiLCIgICAgICAzNjM4RCAgIn0 = " => array(
                "name" => 'СТИН - сервис',
                "clients" => array(
                    "eyJCIiwiMCIsIjAiLCIxNzIiLCIwIiwiMCIsIiAgICAgICA1MDRTICAifQ == " => array(
                    "name" => 'СТИН - сервис!!!',
                    "realizes_date" => array(
                            "27.06 .2022 - 03.07 .2022" => array(
                                "summ" => 78331.66,
                                "calls" => 1
                            ),
                            "04.07 .2022 - 10.07 .2022" => array(
                                "summ" => 403287.96,
                                "calls" => 0
                            ),
                            "11.07 .2022 - 17.07 .2022" => array(
                                "summ" => 418194.42,
                                "calls" => 0
                            )

                            ),

                         "dz" => 638583.15,
                         "tdz" => 0,
                         "pdz" => 638583.15,
                         "summ" => 999200.63,
                         "calls" => 1
                     )

                 )
             )
         )
                        ),
                        "eyJCIiwiMCIsIjAiLCI1NTYiLCIwIiwiMCIsIiAgICAgICAgMTdEICAiee == " => array(
                            "name" => 'Коммерческий отдел1',
                            "groups" => array(
                                "eyJCIiwiMCIsIjAiLCI5NjMzIiwiMCIsIjAiLCIgICAgICAzNjM4RCAgIn0 = " => array(
                                   "name" => 'СТИН - сервис1',
                                   "clients" => array(
                                       "eyJCIiwiMCIsIjAiLCIxNzIiLCIwIiwiMCIsIiAgICAgICA1MDRTICAifQ == " => array(
                                       "name" => 'СТИН - сервис!1!!',
                                       "realizes_date" => array(
                                               "27.06 .2022 - 03.07 .2022" => array(
                                                   "summ" => 78331.66,
                                                   "calls" => 0
                                               ),
                                               "04.07 .2022 - 10.07 .2022" => array(
                                                   "summ" => 403287.96,
                                                   "calls" => 1
                                               ),
                                               "11.07 .2022 - 17.07 .2022" => array(
                                                   "summ" => 418194.42,
                                                   "calls" => 1
                                               )
                   
                                               ),
                   
                                            "dz" => 638583.15,
                                            "tdz" => 0,
                                            "pdz" => 638583.15,
                                            "summ" => 999200.63,
                                            "calls" => 2
                                            ),
                                            "eyJCIiwiMCIsIjAiLCIxNzIiLCIwIiwiMCIsIiAgICAgICA1hMDRTICAifQ == " => array(
                                                "name" => 'СТИН - сервис!1!!',
                                                "realizes_date" => array(
                                                        "27.06 .2022 - 03.07 .2022" => array(
                                                            "summ" => 78331.66,
                                                            "calls" => 0
                                                        ),
                                                        "04.07 .2022 - 10.07 .2022" => array(
                                                            "summ" => 403287.96,
                                                            "calls" => 1
                                                        ),
                                                        "11.07 .2022 - 17.07 .2022" => array(
                                                            "summ" => 418194.42,
                                                            "calls" => 1
                                                        )
                            
                                                        ),
                            
                                                     "dz" => 638583.15,
                                                     "tdz" => 0,
                                                     "pdz" => 638583.15,
                                                     "summ" => 999200.63,
                                                     "calls" => 2
                                                 )
                   
                                    )
                                )
                            )
                        )
 );

 require_once('Classes/PHPExcel.php');
 $phpexcel = new PHPExcel();


 # код для массивов

 $letters = array(
	'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R',
	'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'
);
 $rows = array('Реализ.:', 'Конт. тел.:', 'Задач:', 'Событий:');

 $i = 1;
 $j = 4;
 $p = 1;
 $phpexcel->getActiveSheet()->setShowSummaryBelow(false);


 foreach ($arr as $key => $index){
    $numOfRows = 0;
    $countClients = 0;
    $countGroups = 0;
    $menedger = $index['name'];
    $phpexcel->getActiveSheet()->setCellValueExplicit("A$i", "Менеджер: $menedger", PHPExcel_Cell_DataType::TYPE_STRING);
    for ($k = 0; $k < 4; $k++) {
        $row = $rows[$k];
        $num = $i + $k + 1;
        $phpexcel->getActiveSheet()->setCellValueExplicit("D$num", "$row", PHPExcel_Cell_DataType::TYPE_STRING);
        $phpexcel->getActiveSheet()->getRowDimension($num)->setOutlineLevel(1);
        $phpexcel->getActiveSheet()->getRowDimension($num)->setVisible(true);
    }
    $numOfRows += 1;
    foreach ($index['groups'] as $managerArray){
        $group = $managerArray['name'];
        $i = $i + 5;
        $phpexcel->getActiveSheet()->setCellValueExplicit("A$i", "Группа: $group", PHPExcel_Cell_DataType::TYPE_STRING);
        $phpexcel->getActiveSheet()->getRowDimension($i)->setOutlineLevel(1);
        $phpexcel->getActiveSheet()->getRowDimension($i)->setVisible(true);
        for ($k = 0; $k < 4; $k++) {
            $row = $rows[$k];
            $num1 = $i + $k + 1;
            $phpexcel->getActiveSheet()->setCellValueExplicit("D$num1", "$row", PHPExcel_Cell_DataType::TYPE_STRING);
            $phpexcel->getActiveSheet()->getRowDimension($num1)->setOutlineLevel(2);
            $phpexcel->getActiveSheet()->getRowDimension($num1)->setVisible(true);
        }
        $numOfRows += 1;
        foreach ($managerArray['clients'] as $clientArray){
            $numOfRows += 1;
            $i = $i + 5;
            $client = $clientArray['name'];
            $phpexcel->getActiveSheet()->setCellValueExplicit("A$i", "Клиент: $client", PHPExcel_Cell_DataType::TYPE_STRING);
            $dz = $clientArray['dz'];
            $pdz = $clientArray['pdz'];
            $tdz = $clientArray['tdz'];
            $summa = $clientArray['summ'];
            $sumCalls = $clientArray['calls'];
            for ($n = 0; $n < $numOfRows; $n++){
                $t = $i + 1 - 5*$n;
                $phpexcel->getActiveSheet()->mergeCells("B$t:C$t");
                $t = $i - 5*$n + 2;
                $phpexcel->getActiveSheet()->mergeCells("B$t:C$t");
                $t = $i - 5*$n + 1;
                $phpexcel->getActiveSheet()->setCellValueExplicit("B$t", "ДЗ", PHPExcel_Cell_DataType::TYPE_STRING);
                $t = $i - 5*$n + 2;
                $phpexcel->getActiveSheet()->setCellValueExplicit("B$t", "$dz", PHPExcel_Cell_DataType::TYPE_STRING);
                $t = $i - 5*$n + 3;
                $phpexcel->getActiveSheet()->setCellValueExplicit("B$t", "ТДЗ", PHPExcel_Cell_DataType::TYPE_STRING);
                $phpexcel->getActiveSheet()->setCellValueExplicit("C$t", "ПДЗ", PHPExcel_Cell_DataType::TYPE_STRING);
                $t = $i - 5*$n + 4;
                $phpexcel->getActiveSheet()->setCellValueExplicit("B$t", $tdz, PHPExcel_Cell_DataType::TYPE_STRING);
                $phpexcel->getActiveSheet()->setCellValueExplicit("C$t", $pdz, PHPExcel_Cell_DataType::TYPE_STRING);
            }
            $j = 4;
            foreach ($clientArray['realizes_date'] as $date => $summ){
                $phpexcel->getActiveSheet()->setCellValueExplicit($letters[$j].$p, "$date", PHPExcel_Cell_DataType::TYPE_STRING);
                $calls = $summ['calls'];
                $ans = $p + 1;
                $ans3 = $p + 2;
                $phpexcel->getActiveSheet()->setCellValueExplicit($letters[$j].$ans, $summ['summ'], PHPExcel_Cell_DataType::TYPE_STRING);
                $phpexcel->getActiveSheet()->setCellValueExplicit($letters[$j].$ans3, $calls, PHPExcel_Cell_DataType::TYPE_STRING);
                $j++;
            }
            $phpexcel->getActiveSheet()->getRowDimension($i)->setOutlineLevel(2);
            $phpexcel->getActiveSheet()->getRowDimension($i)->setVisible(true);
            for ($k = 0; $k < 4; $k++) {
                $row1 = $rows[$k];
                $num2 = $i + $k + 1;
                $phpexcel->getActiveSheet()->setCellValueExplicit("D$num2", "$row1", PHPExcel_Cell_DataType::TYPE_STRING);
                $phpexcel->getActiveSheet()->getRowDimension($num2)->setOutlineLevel(3);
                $phpexcel->getActiveSheet()->getRowDimension($num2)->setVisible(true);
            }
            $phpexcel->getActiveSheet()->setCellValueExplicit($letters[$j].$p, "Итого:", PHPExcel_Cell_DataType::TYPE_STRING);
            $ans = $p + 1;
            $phpexcel->getActiveSheet()->setCellValueExplicit($letters[$j].$ans, $summa, PHPExcel_Cell_DataType::TYPE_STRING);
            $ans = $p + 2;
            $phpexcel->getActiveSheet()->setCellValueExplicit($letters[$j].$ans, $sumCalls, PHPExcel_Cell_DataType::TYPE_STRING);
        }
    }
    $i = $i + 5;
    $p = $i;
 }


 $page = $phpexcel->setActiveSheetIndex();
 $page->setTitle("static-detail");
 $objWriter = PHPExcel_IOFactory::createWriter($phpexcel, 'Excel2007');
 $filename = "static-detail.xlsx";
 if( file_exists($filename) ){
    unlink($filename);
 }
 $objWriter->save($filename);
 ?>
