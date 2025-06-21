<?php
require_once dirname(dirname(dirname(dirname(__FILE__)))) . '/config.php';

define('SP_TEACHER_NO_EDITING', 18);
define('SP_STUDENT', 9);
define('SP_PRESIDENTE', 10);
define('SP_VICEPRESIDENTE', 11);
define('SP_FINANZAS', 12);
define('SP_LOGISTICA', 13);
define('SP_TALENTO_HUMANO', 14);

global $DB, $CFG;

require_once __DIR__ . '/componente/ws_data.php';
require_once __DIR__ . '/componente/db_data.php';
require_once __DIR__ . '/componente/suspended_data.php';
require_once __DIR__ . '/componente/firstpage.php';
require_once __DIR__ . '/componente/cellColor.php';

// $time_start = microtime(true);
$categoryid  = $DB->get_record('config', ['name' => 'wssplashcategoryid'])->value;
$role_config = [SP_STUDENT];

$columnos_excel = [
    'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T',
    'U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL',
    'AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB','BC',
    'BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT',
    'BU','BV','BW','BX','BY','BZ','CA','CB','CC','CD','CE','CF','CG','CH','CI','CJ','CK',
    'CL','CM','CN'
];

$output = [];

$courses = $DB->get_records('course', ['category' => $categoryid, 'visible' => 1], '', 'id,fullname, idnumber,shortname as aula,visible');

function cmp($a, $b) {
    return strcmp($a->fullname, $b->fullname);
}
usort($courses, "cmp");

$second_courses = $DB->get_records('course', ['category' => $categoryid], '', 'id,fullname,shortname as aula,visible');
usort($second_courses, "cmp");

$glob_roles = [];

foreach ($courses as $key => $value) {
    $tmp_data = explode('-', $value->fullname);
    unset($value->fullname);

    $pos = strpos($tmp_data[0], 'MADRE DE DIOS');
    if ($pos === false) {
        $split_name = explode(' DE ', $tmp_data[0]);
        $value->departamento = $split_name[count($split_name) - 1];
    } else {
        $value->departamento = 'MADRE DE DIOS';
    }

    $value->name         = $tmp_data[0];
    $value->equipo       = isset($tmp_data[3]) ? $tmp_data[3] : '';
    $value->categoria    = isset($tmp_data[4]) ? $tmp_data[4] : '';

    $sql_roles = "SELECT concat(u.id,r.id,ra.id) as ignoreid, u.id as userid, u.firstname, u.lastname, u.email, u.username as dni, u.address, r.id as roleid, r.name as rolfulname, r.shortname as rolename
                  FROM mdl_user u
                  JOIN mdl_user_enrolments ue ON ue.userid = u.id
                  JOIN mdl_enrol e ON e.id = ue.enrolid
                  JOIN mdl_role_assignments ra ON ra.userid = u.id
                  JOIN mdl_context ct ON ct.id = ra.contextid AND ct.contextlevel = 50
                  JOIN mdl_course c ON c.id = ct.instanceid AND e.courseid = c.id
                  JOIN mdl_role r ON r.id = ra.roleid
                  WHERE  c.id = " . $value->id . "
                  AND e.status = 0 AND u.suspended = 0 AND u.deleted = 0
                  AND ue.status = 0 ORDER BY u.lastname ASC";

    $sql = $DB->get_records_sql($sql_roles);

    $value->users = [];
    $sql_tot_splash = "SELECT concat(u.id,r.id,ra.id) as ignoreid, u.id as userid, u.firstname, u.lastname, u.username as dni, u.email, u.address, r.id as roleid, r.shortname as rolename, u.city as departamento
                  FROM mdl_user u
                  JOIN mdl_user_enrolments ue ON ue.userid = u.id
                  JOIN mdl_enrol e ON e.id = ue.enrolid
                  JOIN mdl_role_assignments ra ON ra.userid = u.id
                  JOIN mdl_context ct ON ct.id = ra.contextid AND ct.contextlevel = 50
                  JOIN mdl_course c ON c.id = ct.instanceid AND e.courseid = c.id
                  JOIN mdl_role r ON r.id = ra.roleid AND r.id = " . SP_STUDENT . "
                  WHERE  c.id = " . $value->id . "
                  AND e.status = 0 AND u.suspended = 0 AND u.deleted = 0
                  AND ue.status = 0 ORDER BY u.lastname ASC";
    $tot_splash = $DB->get_records_sql($sql_tot_splash);

    $prof = 0;
    $roles = [];

    foreach ($sql as $ke => $va) {
        if ($va->userid == $prof) {}
        if (SP_TEACHER_NO_EDITING == $va->roleid) {
            $prof = $va->userid;
        }
        $roles[$va->roleid]        = ['name' => $va->rolename, 'valeu' => ''];
        $glob_roles[$va->rolename] = '';
    }

    $wp_data   = [];
    $tmp_keys  = array_keys($sql);
    $tmp_key_c = '';
    $repetidos = [];

    foreach ($sql as $ke => $va) {
        $irse = false;
        foreach ($repetidos as $repe) {
            if ($ke == $repe) $irse = true;
        }
        if ($irse) continue;
        $name_roles = $roles;

        if (SP_TEACHER_NO_EDITING != $va->roleid) {
            $name_roles[$va->roleid]['valeu'] = '1';
        }

        if (SP_TEACHER_NO_EDITING == $va->roleid) {
            $prof                             = $va->userid;
            $va->spl                          = count($tot_splash);
            $name_roles[$va->roleid]['valeu'] = '1';
            $tmp_key_c = $ke;
        } else if ($prof != $va->userid) {
            $va->spl = '';
        }

        foreach ($sql as $ksql => $vsql) {
            if ($va->userid == $vsql->userid && $va->ignoreid != $vsql->ignoreid) {
                if (SP_TEACHER_NO_EDITING != $vsql->roleid) {
                    $name_roles[$vsql->roleid]['valeu'] = '1';
                }
                if (SP_TEACHER_NO_EDITING == $vsql->roleid) {
                    $prof                                       = $vsql->userid;
                    $sql[$ke]->spl                              = count($tot_splash);
                    $va->spl                                    = count($tot_splash);
                    $name_roles[SP_TEACHER_NO_EDITING]['valeu'] = '1';
                    $tmp_key_c        = $ke;
                    $sql[$ke]->roleid = SP_TEACHER_NO_EDITING;
                    $va->roleid       = SP_TEACHER_NO_EDITING;
                } else if ($prof != $va->userid) {
                    $va->spl = '';
                }
                $repetidos[] = $ksql;
                unset($sql[$ksql]);
            }
        }

        $va->roles = array_values($name_roles);

        $sql_login = "SELECT c.id, c.fullname as virtualroom, COUNT(usr.username) as total
            FROM mdl_course c
            INNER JOIN mdl_context cx ON c.id = cx.instanceid
            INNER JOIN mdl_role_assignments ra ON cx.id = ra.contextid
            INNER JOIN mdl_user usr ON ra.userid = usr.id
            INNER JOIN mdl_role r ON ra.roleid = r.id
            INNER JOIN mdl_logstore_standard_log log ON log.userid = usr.id
            WHERE cx.contextlevel = '50'
            AND c.category='" . $categoryid . "'
            AND c.id = " . $value->id . "
            AND ra.roleid <> 'null'
            AND log.action='loggedin'
            GROUP BY c.fullname
            ORDER BY c.fullname";
        $totLogins = $DB->get_record_sql($sql_login);
        $va->totlogins = is_object($totLogins) ? $totLogins->total : '0';

        $data    = [];
        $puntaje = 0;
        $nsemana = 17;
        for ($i = 2; $i < $nsemana; $i++) {
            $exist_data = $DB->get_record('local_wsplashdata', ['courseid' => $value->id, 'semana' => $i]);
            if (is_object($exist_data)) {
                $tmp_data = json_decode($exist_data->data);
                if ('Tarea' == $tmp_data[0]->tipo && $tmp_data[0]->data != []) {
                    $tmp_data[0]->archivos = db_data($tmp_data[0]->data[0]);
                }
            } else {
                $tmp_data = ws_data($value->id, $i, SP_STUDENT);
            }

            if ([] == $tmp_data || !isset($tmp_data[0]->section_name)) {
                if($va->roleid == SP_STUDENT){
                    $va->totlogins    = '';
                    $va->puntajetotal = '';
                }
                continue;
            }

            if (count($tmp_data) > 1) {
                foreach ($tmp_data as $q => $b) {
                    if($b->archivos > 0){
                        $tmp_data[array_keys($tmp_data)[0]] = $tmp_data[$q];
                    }
                }
                unset($tmp_data[array_keys($tmp_data)[1]]);
            }

            $data['Semana ' . $tmp_data[0]->section_name] = $tmp_data;

            if (!isset($tmp_data[0]->section_name)) {
                echo "<pre>";
                print_r($value);
                print_r(ws_data($value->id, $i, SP_STUDENT));
                print_r($tmp_data);
                echo "</pre>";die();
            }

            foreach ($data['Semana ' . $tmp_data[0]->section_name] as $k => $v) {
                $puntaje += ('' == $v->puntaje) ? 0 : $v->puntaje;
            }

            if (count($data['Semana ' . $tmp_data[0]->section_name]) > 1) {
                foreach ($data['Semana ' . $tmp_data[0]->section_name] as $k => $v) {
                    $data['Semana ' . $tmp_data[0]->section_name . ' ' . $k] = [$v];
                }
                unset($data['Semana ' . $tmp_data[0]->section_name]);
            }
        }

        $va->data         = array_values($data);
        $va->puntajetotal = $puntaje;

        if (SP_TEACHER_NO_EDITING != $va->roleid) {
            $va->totlogins    = '';
            $va->puntajetotal = '';
            if ([] == $data) continue;
            foreach ($va->roles as $kie => $vue) {
                if ('s' == $vue['name']) {
                    $va->roles[$kie]['valeu'] = 1;
                }
            }
            foreach ($va->data as $keee => $a_vaaa) {
                foreach ($a_vaaa as $kee => $vaa) {
                    $a_vaaa[$kee]->activity_name = '';
                    $a_vaaa[$kee]->section_name  = '';
                    $a_vaaa[$kee]->puntaje       = '';
                    $a_vaaa[$kee]->tipo          = '';
                    $a_vaaa[$kee]->archivos      = '';
                }
            }
        }
    }
    $sql_wi = [];
    if ([] != $sql && '' != $tmp_key_c) {
        $sql_wi['temporal'] = $sql[$tmp_key_c];
        unset($sql[$tmp_key_c]);
        array_unshift($sql, $sql_wi['temporal']);
    }
    $value->users[] = array_values($sql);
}

$output = $courses;

$max = 0;
foreach ($output as $value) {
    if (isset($value->users[0][0]->data) && count($value->users[0][0]->data) > $max) {
        $max = count($value->users[0][0]->data);
    }
}

// $time_end = microtime(true);
// $execution_time = ($time_end - $time_start)/60;

$tmp_roles  = [];
$tmp_roles2 = [];
foreach ($glob_roles as $kkk => $vvv) {
    $tmp_roles2[$kkk] = '1';
}
unset($tmp_roles2['s']);
foreach (['c', 'p', 'vp', 'gaf', 'glo', 'gth', 'best'] as $role) {
    if (isset($glob_roles[$role])) {
        unset($glob_roles[$role]);
        unset($tmp_roles2[$role]);
        $tmp_roles[$role] = '';
    }
}
$glob_roles = ['s' => ''] + $tmp_roles2 + $tmp_roles;

require_once '../classes/PHPExcel.php';

$titulo = "Inscritos y Puntajes";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="' . $titulo . '.xlsx"');
header('Cache-Control: max-age=0');

$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Jair Revilla")
    ->setLastModifiedBy("Jair Revilla")
    ->setTitle($titulo)
    ->setSubject($titulo)
    ->setDescription($titulo)
    ->setKeywords($titulo)
    ->setCategory("Reporte excel");

$i = 0;
$objWorkSheet = $objPHPExcel->createSheet($i);

firstpage($objWorkSheet, $glob_roles, $columnos_excel, $output, $max, 'Activos', $tmp_roles2);

$i++;
$output = suspended_data($second_courses, $categoryid);
$objWorkSheet = $objPHPExcel->createSheet($i);

firstpage($objWorkSheet, $output[1], $columnos_excel, $output[0], $max, 'Suspendidos', $tmp_roles2, true);

$objPHPExcel->setActiveSheetIndexByName('Worksheet');
$sheetIndex = $objPHPExcel->getActiveSheetIndex();
$objPHPExcel->removeSheetByIndex($sheetIndex);

foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
    $objPHPExcel->setActiveSheetIndex($objPHPExcel->getIndex($worksheet));
    $sheet        = $objPHPExcel->getActiveSheet();
    $cellIterator = $sheet->getRowIterator()->current()->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(true);
    foreach ($cellIterator as $cell) {
        $sheet->getColumnDimension($cell->getColumn())->setAutoSize(true);
    }
}

$objPHPExcel->setActiveSheetIndex(0);

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
die();