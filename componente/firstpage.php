<?php
function firstpage($objPHPExcel, $glob_roles, $columnos_excel, $output, $max, $nombreHoja, $tmp_roles2, $suspend = false) {
    require_once '../classes/PHPExcel.php';

    $header_excel   = 0;
    $objPHPExcel
        ->setCellValue($columnos_excel[$header_excel++] . '2', '#')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Departamento')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Nombre del Colegio')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Aula')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Equipo')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Categoría')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Apellido')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Nombre')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'DNI')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Email')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Rol');

    $excel_roles = [];
    foreach ($glob_roles as $key => $value) {
        $excel_roles[$key] = $columnos_excel[$header_excel];
        $objPHPExcel->setCellValue($columnos_excel[$header_excel++] . '2', strtoupper($key));
    }
    $objPHPExcel->setCellValue($columnos_excel[$header_excel++] . '2', 'LOG');

    for ($i = 2; $i <= $max; $i++) {
        $first_column  = $header_excel++;
        $second_column = $header_excel++;
        $objPHPExcel
            ->setCellValue($columnos_excel[$first_column] . '1', 'SEM ' . $i)
            ->mergeCells($columnos_excel[$first_column] . '1:' . $columnos_excel[$second_column] . '1')
            ->setCellValue($columnos_excel[$first_column] . '2', 'PUNT')
            ->setCellValue($columnos_excel[$second_column] . '2', (10 == $i || 12 == $i || 16 == $i) ? 'RPTA' : 'ARCH');
    }

    $first_column  = $header_excel++;
    $second_column = $header_excel++;

    $objPHPExcel
        ->setCellValue($columnos_excel[$first_column] . '1', 'SEM 16')
        ->mergeCells($columnos_excel[$first_column] . '1:' . $columnos_excel[$second_column] . '1')
        ->setCellValue($columnos_excel[$first_column] . '2', 'PUNT')
        ->setCellValue($columnos_excel[$second_column] . '2', 'RPTA');

    $objPHPExcel
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Audit 1')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Audit 2')
        ->setCellValue($columnos_excel[$header_excel] . '2', 'TOTAL ')
        ->setCellValue('A1', 'Datos al ' . date("d/m/Y g:i a", time()));

    $objPHPExcel
        ->mergeCells('A1:C1');
    $objPHPExcel->getStyle('A1')->getFont()->setBold(true);
    $objPHPExcel->getStyle('A2:' . $columnos_excel[$header_excel] . '2')->getAlignment()->setHorizontal(
        PHPExcel_Style_Alignment::HORIZONTAL_CENTER
    );
    $objPHPExcel->getStyle('A1:' . $columnos_excel[$header_excel] . '1')->getAlignment()->setHorizontal(
        PHPExcel_Style_Alignment::HORIZONTAL_CENTER
    );
    $objPHPExcel
        ->getStyle('A2:' . $columnos_excel[$header_excel] . '2')
        ->getFill()
        ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        ->getStartColor()
        ->setRGB('95B3D7');

    $actual    = 3;
    $index_con = 1;
    foreach ($output as $key => $value) {
        $tmp_index = $index_con++;
        foreach ($value->users[0] as $val) {
            $column_data = 0;
            $objPHPExcel
                ->setCellValue($columnos_excel[$column_data++] . $actual, $tmp_index)
                ->setCellValue($columnos_excel[$column_data++] . $actual, $value->departamento)
                ->setCellValue($columnos_excel[$column_data++] . $actual, $value->name)
                ->setCellValue($columnos_excel[$column_data++] . $actual, $value->aula)
                ->setCellValue($columnos_excel[$column_data++] . $actual, $value->equipo)
                ->setCellValue($columnos_excel[$column_data++] . $actual, $value->categoria)
                ->setCellValue($columnos_excel[$column_data++] . $actual, mb_strtoupper($val->lastname, 'utf-8'))
                ->setCellValue($columnos_excel[$column_data++] . $actual, mb_strtoupper($val->firstname, 'utf-8'))
                ->setCellValue($columnos_excel[$column_data++] . $actual, (string) $val->dni)
                ->setCellValue($columnos_excel[$column_data++] . $actual, $val->email)
                ->setCellValue($columnos_excel[$column_data++] . $actual, $val->rolfulname);

            foreach ($val->roles as $rolee) {
                $objPHPExcel->setCellValue($excel_roles[$rolee['name']] . $actual, $rolee['valeu']);
            }
            $column_data += count($excel_roles);

            $objPHPExcel->setCellValue($columnos_excel[$column_data++] . $actual, $val->totlogins);

            foreach ($val->data as $puntajee) {
                $objPHPExcel
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $puntajee[0]->puntaje)
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $puntajee[0]->archivos);
            }

            $objPHPExcel
                ->setCellValue($columnos_excel[$column_data++] . $actual, "") // Audit 1 (puedes implementar si tienes datos)
                ->setCellValue($columnos_excel[$column_data++] . $actual, "") // Audit 2 (puedes implementar si tienes datos)
                ->setCellValue($columnos_excel[$column_data++] . $actual, $val->puntajetotal);

            $actual++;
        }
    }
    $objPHPExcel->freezePaneByColumnAndRow(0, 3);
    $objPHPExcel->setTitle($nombreHoja);
}