<?php
function cellColor($cells, $color, $objPHPExcel) {
    $objPHPExcel->getStyle($cells)->getFill()->applyFromArray([
        'type'       => PHPExcel_Style_Fill::FILL_SOLID,
        'startcolor' => [
            'rgb' => $color,
        ],
    ]);
}