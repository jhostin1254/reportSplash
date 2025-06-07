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



$categoryid  = $DB->get_record('config', ['name' => 'wssplashcategoryid'])->value;
$role_config = [SP_STUDENT];

$columnos_excel = ['A',
    'B',
    'C',
    'D',
    'E',
    'F',
    'G',
    'H',
    'I',
    'J',
    'K',
    'L',
    'M',
    'N',
    'O',
    'P',
    'Q',
    'R',
    'S',
    'T',
    'U',
    'V',
    'W',
    'X',
    'Y',
    'Z',
    'AA',
    'AB',
    'AC',
    'AD',
    'AE',
    'AF',
    'AG',
    'AH',
    'AI',
    'AJ',
    'AK',
    'AL',
    'AM',
    'AN',
    'AO',
    'AP',
    'AQ',
    'AR',
    'AS',
    'AT',
    'AU',
    'AV',
    'AW',
    'AX',
    'AY',
    'AZ',
    'BA',
    'BB',
    'BC',
    'BD',
    'BE',
    'BF',
    'BG',
    'BH',
    'BI',
    'BJ',
    'BK',
    'BL',
    'BM',
    'BN',
    'BO',
    'BP',
    'BQ',
    'BR',
    'BS',
    'BT',
    'BU',
    'BV',
    'BW',
    'BX',
    'BY',
    'BZ',
    'CA',
    'CB',
    'CC',
    'CD',
    'CE',
    'CF',
    'CG',
    'CH',
    'CI',
    'CJ',
    'CK',
    'CL',
    'CM',
    'CN'];

$output = [];

$courses = $DB->get_records('course', ['category' => $categoryid, 'visible' => 1], '', 'id,fullname, idnumber,shortname as aula,visible');
$second_courses = $DB->get_records('course', ['category' => $categoryid], '', 'id,fullname,shortname as aula,visible');

function cmp($a, $b)
{
    return strcmp($a->fullname, $b->fullname);
}

usort($courses, "cmp");
usort($second_courses, "cmp");

$glob_roles = [];

//por wue no en vez de wue cada vez que se llama a la base de datos no se guarda en un array y se usa ese array
foreach ($courses as $key => $value) {
    // Procesamiento inicial del nombre del curso
    $tmp_data = explode('-', $value->fullname);
    unset($value->fullname);

    // Extraer el departamento
    $split_name = explode(' DE ', $tmp_data[0]);
    $value->departamento = end($split_name);

    // Asignación de valores básicos
    $value->name = $tmp_data[0] ?? '';
    $value->equipo = $tmp_data[3] ?? '';
    $value->categoria = $tmp_data[4] ?? '';

    // Consulta para usuarios y roles
    $sql_roles = "SELECT concat(u.id,r.id,ra.id) as ignoreid, u.id as userid, 
                 u.firstname, u.lastname, u.email, u.username as dni, u.address, 
                 r.id as roleid, r.name as rolfulname, r.shortname as rolename, 
                 u.phone2, u.city as departamento, c.id as courseid
              FROM mdl_user u
              JOIN mdl_user_enrolments ue ON ue.userid = u.id
              JOIN mdl_enrol e ON e.id = ue.enrolid
              JOIN mdl_role_assignments ra ON ra.userid = u.id
              JOIN mdl_context ct ON ct.id = ra.contextid AND ct.contextlevel = 50
              JOIN mdl_course c ON c.id = ct.instanceid AND e.courseid = c.id
              JOIN mdl_role r ON r.id = ra.roleid
              WHERE c.id = :courseid
              AND e.status = 0 AND u.suspended = 0 AND u.deleted = 0
              AND ue.status = 0 
              ORDER BY u.lastname ASC";

    $sql = $DB->get_records_sql($sql_roles, ['courseid' => $value->id]);

    // Consulta para splashers (estudiantes)
    $sql_tot_splash = "SELECT concat(u.id,r.id,ra.id) as ignoreid, u.id as userid, 
                      u.firstname, u.lastname, u.username as dni, u.email, u.address, 
                      r.id as roleid, r.shortname as rolename, u.city as departamento, 
                      c.id as courseid
                   FROM mdl_user u
                   JOIN mdl_user_enrolments ue ON ue.userid = u.id
                   JOIN mdl_enrol e ON e.id = ue.enrolid
                   JOIN mdl_role_assignments ra ON ra.userid = u.id
                   JOIN mdl_context ct ON ct.id = ra.contextid AND ct.contextlevel = 50
                   JOIN mdl_course c ON c.id = ct.instanceid AND e.courseid = c.id
                   JOIN mdl_role r ON r.id = ra.roleid AND r.id = :studentrole
                   WHERE c.id = :courseid
                   AND e.status = 0 AND u.suspended = 0 AND u.deleted = 0
                   AND ue.status = 0 
                   ORDER BY u.lastname ASC";

    $tot_splash = $DB->get_records_sql($sql_tot_splash, [
        'courseid' => $value->id,
        'studentrole' => SP_STUDENT
    ]);

    $prof = 0;
    $roles = [];
    $glob_roles = [];
    $processed_ids = [];
    $teacher_data = null;

    // Procesar roles y identificar profesor
    foreach ($sql as $ke => $va) {
        if (SP_TEACHER_NO_EDITING == $va->roleid) {
            $prof = $va->userid;
            $va->spl = count($tot_splash);
            $teacher_data = $va;
        }
        $roles[$va->roleid] = ['name' => $va->rolename, 'valeu' => ''];
        $glob_roles[$va->rolename] = '';
    }

    // Procesar usuarios
    $wp_data = [];
    $repetidos = [];
    $final_users = [];

    foreach ($sql as $ke => $va) {
        if (in_array($va->ignoreid, $processed_ids)) {
            continue;
        }
        $processed_ids[] = $va->ignoreid;

        $name_roles = $roles;

        if (SP_TEACHER_NO_EDITING == $va->roleid) {
            $name_roles[$va->roleid]['valeu'] = '1';
            $va->spl = count($tot_splash);
        } else {
            $va->spl = '';
        }

        // Manejar usuarios con múltiples roles
        foreach ($sql as $ksql => $vsql) {
            if ($va->userid == $vsql->userid && $va->ignoreid != $vsql->ignoreid) {
                if (SP_TEACHER_NO_EDITING == $vsql->roleid) {
                    $prof = $vsql->userid;
                    $va->spl = count($tot_splash);
                    $name_roles[SP_TEACHER_NO_EDITING]['valeu'] = '1';
                    $va->roleid = SP_TEACHER_NO_EDITING;
                }
                $repetidos[] = $ksql;
                unset($sql[$ksql]);
            }
        }

        $va->roles = array_values($name_roles);

        // Obtener total de logins
        $sql_login = "SELECT COUNT(usr.username) as total
            FROM mdl_course c
            INNER JOIN mdl_context cx ON c.id = cx.instanceid
            INNER JOIN mdl_role_assignments ra ON cx.id = ra.contextid
            INNER JOIN mdl_user usr ON ra.userid = usr.id
            INNER JOIN mdl_role r ON ra.roleid = r.id
            INNER JOIN mdl_logstore_standard_log log ON log.userid = usr.id
            WHERE cx.contextlevel = '50'
            AND c.category = :categoryid
            AND c.id = :courseid
            AND ra.roleid <> 'null'
            AND log.action = 'loggedin'";

        $totLogins = $DB->get_field_sql($sql_login, [
            'categoryid' => $categoryid,
            'courseid' => $value->id
        ]);
        $va->totlogins = $totLogins ?: '0';

        // Procesar semanas
        $data = [];
        $puntaje = 0;
        $nsemana = 17; // $DB->get_record('config', ['name' => 'reportpointscantsem'])->value;

        for ($i = 2; $i < $nsemana; $i++) {
            $exist_data = $DB->get_record('local_wsplashdata', [
                'courseid' => $value->id,
                'semana' => $i
            ]);

            $tmp_data = $exist_data ? 
                json_decode($exist_data->data) : 
                ws_data($value->id, $i, SP_STUDENT);

            if (empty($tmp_data) || !isset($tmp_data[0]->section_name)) {
                if ($va->roleid == SP_STUDENT) {
                    $va->totlogins = '';
                    $va->puntajetotal = '';
                }
                continue;
            }

            // Seleccionar datos principales
            if (count($tmp_data) > 1) {
                foreach ($tmp_data as $q => $b) {
                    if ($b->archivos > 0) {
                        $tmp_data[0] = $b;
                        break;
                    }
                }
                $tmp_data = [$tmp_data[0]];
            }

            $semana_key = 'Semana ' . $tmp_data[0]->section_name;
            $data[$semana_key] = $tmp_data;

            // Calcular puntaje
            foreach ($data[$semana_key] as $v) {
                $puntaje += $v->puntaje ?? 0;
            }

            // Manejar múltiples actividades por semana
            if (count($data[$semana_key]) > 1) {
                foreach ($data[$semana_key] as $k => $v) {
                    $data["{$semana_key} {$k}"] = [$v];
                }
                unset($data[$semana_key]);
            }
        }

        $va->data = array_values($data);
        $va->puntajetotal = $puntaje;

        // Limpiar datos para no profesores
        if (SP_TEACHER_NO_EDITING != $va->roleid) {
            $va->totlogins = '';
            $va->puntajetotal = '';

            if (empty($data)) {
                continue;
            }

            foreach ($va->roles as $kie => $vue) {
                if ($vue['name'] == 's') {
                    $va->roles[$kie]['valeu'] = 1;
                }
            }

            foreach ($va->data as $keee => $a_vaaa) {
                foreach ($a_vaaa as $kee => $vaa) {
                    $a_vaaa[$kee]->activity_name = '';
                    $a_vaaa[$kee]->section_name = '';
                    $a_vaaa[$kee]->puntaje = '';
                    $a_vaaa[$kee]->tipo = '';
                    $a_vaaa[$kee]->archivos = '';
                }
            }
        }

        $final_users[] = $va;
    }

    // Colocar al profesor primero si existe
    if ($teacher_data) {
        array_unshift($final_users, $teacher_data);
    }

    $value->users = [array_values($final_users)];
}
$output = $courses;

$max = 0;
foreach ($output as $value) {
    if (isset($value->users[0][0]->data) && count($value->users[0][0]->data) > $max) {
        $max = count($value->users[0][0]->data);
    }
}

$time_end = microtime(true);

//dividing with 60 will give the execution time in minutes otherwise seconds
$execution_time = ($time_end - $time_start)/60;

//execution time of the script
//echo '<b>Total Execution Time:</b> '.$execution_time.' Mins';

//echo "<pre>";
//print_r($courses);
//echo "</pre>";die();

$tmp_roles  = [];
$tmp_roles2 = [];

foreach ($glob_roles as $kkk => $vvv) {
    $tmp_roles2[$kkk] = '1';
}
unset($tmp_roles2['s']);

if (isset($glob_roles['c'])) {
    unset($glob_roles['c']);
    unset($tmp_roles2['c']);
    $tmp_roles['c'] = '';
}

if (isset($glob_roles['p'])) {
    unset($glob_roles['p']);
    unset($tmp_roles2['p']);
    $tmp_roles['p'] = '';
}

if (isset($glob_roles['vp'])) {
    unset($glob_roles['vp']);
    unset($tmp_roles2['vp']);
    $tmp_roles['vp'] = '';
}

if (isset($glob_roles['gaf'])) {
    unset($glob_roles['gaf']);
    unset($tmp_roles2['gaf']);
    $tmp_roles['gaf'] = '';
}

if (isset($glob_roles['glo'])) {
    unset($glob_roles['glo']);
    unset($tmp_roles2['glo']);
    $tmp_roles['glo'] = '';
}

if (isset($glob_roles['gth'])) {
    unset($glob_roles['gth']);
    unset($tmp_roles2['gth']);
    $tmp_roles['gth'] = '';
}

if (isset($glob_roles['best'])) {
    unset($glob_roles['best']);
    unset($tmp_roles2['best']);
    $tmp_roles['best'] = '';
}

$glob_roles = ['s' => ''] + $tmp_roles2 + $tmp_roles;

unset($tmp_roles);
//unset($tmp_roles2);
/*

echo "<pre>";
//print_r($output);
print_r($glob_roles);
echo "</pre>";

echo "<pre>";
//print_r($array_order);
echo "</pre>";

//echo $max . "<br>";

die();

echo "<pre>";
print_r($output);
echo "</pre>";
die();
*/
require_once '../classes/PHPExcel.php';

$titulo = "Inscritos y Puntajes";

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="' . $titulo . '.xlsx"');
header('Cache-Control: max-age=0');
/*
 */
// Se crea el objeto PHPExcel
$objPHPExcel = new PHPExcel();

// Se asignan las propiedades del libro
$objPHPExcel->getProperties()->setCreator("Jair Revilla") // Nombre del autor
    ->setLastModifiedBy("Jair Revilla") //Ultimo usuario que lo modificó
    ->setTitle($titulo) // Titulo
    ->setSubject($titulo) //Asunto
    ->setDescription($titulo) //Descripción
    ->setKeywords($titulo) //Etiquetas
    ->setCategory("Reporte excel"); //Categorias

$i = 0;

/*
echo "<pre>";
print_r($output);
echo "</pre>";
die();

 */
//Profesion
$objWorkSheet = $objPHPExcel->createSheet($i);

firstpage($objWorkSheet, $glob_roles, $columnos_excel, $output, $max, 'Activos', $tmp_roles2);

$i++;
$output = suspended_data($second_courses, $categoryid);
$objWorkSheet = $objPHPExcel->createSheet($i);

firstpage($objWorkSheet, $output[1], $columnos_excel, $output[0], $max, 'Suspendidos', $tmp_roles2,true);

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

/*
echo "<pre>";
print_r($output);
echo "</pre>";

die();
 */

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
die();

function cellColor($cells, $color, $objPHPExcel)
{

    $objPHPExcel->getStyle($cells)->getFill()->applyFromArray([
        'type'       => PHPExcel_Style_Fill::FILL_SOLID,
        'startcolor' => [
            'rgb' => $color,
        ],
    ]);
}

function ws_data($courseid, $section_course, $role_config)
{
    global $DB, $CFG;
    require_once $CFG->libdir . '/adminlib.php';
    require_once $CFG->libdir . '/modinfolib.php';
    require_once $CFG->libdir . '/formslib.php';

    $categoryid     = $courseid;
    $section_course = $section_course;
    $role_config    = [$role_config];

    $role_config = " asg.roleid = " . $role_config[0] . " ";

    if (16 == $section_course) {
        $role_config = " (asg.roleid = " . SP_STUDENT . " OR asg.roleid = " . SP_TEACHER_NO_EDITING . ") ";
    }

    /*
    Retorna los cursos de la categoria elegida y cantidad de alumnos matriculados
     */

    $sql_cursos = "SELECT course.id as Course_id,course.fullname AS Course
     ,context.id AS Context
     , COUNT(course.id) AS Students
     ,category.name
     ,category.path
     FROM {role_assignments} AS asg
     JOIN {context} AS context ON asg.contextid = context.id AND context.contextlevel = 50
     JOIN {user} AS USER ON USER.id = asg.userid
     JOIN {course} AS course ON context.instanceid = course.id
     JOIN {course_categories} AS category ON course.category = category.id
     WHERE " . $role_config . "
     AND course.id =" . $categoryid . "
     GROUP BY course.id
     ORDER BY course.fullname ASC";

    $cursos = $DB->get_records_sql($sql_cursos);

    if (10 == $section_course) {
        $role_config = [10, 11, 12, 13, 14];
    } else if (12 == $section_course) {
        $role_config = [SP_TEACHER_NO_EDITING];
    } else if (13 == $section_course) {
        $role_config = [SP_TEACHER_NO_EDITING];
    } else if (16 == $section_course) {
        $role_config = [SP_TEACHER_NO_EDITING, SP_STUDENT];
    } else {
        $role_config = [SP_STUDENT];
    }

    if ([] == $cursos) {
        //echo 'Verifique la categoria seleccionda';
        $tmp_return          = new stdClass();
        $tmp_return->puntaje = '';
        return [$tmp_return];
    } else {

        foreach ($cursos as $key => $value) {
            if (count(explode('/', $value->path)) > 2) {
                unset($cursos[$key]);
            }
        }

        if ([] == $cursos) {
            $tmp_return          = new stdClass();
            $tmp_return->puntaje = '';
            return [$tmp_return];
        }
    }
    /**
    devuelve datos de las actividades(encuesta y tareas)
     */

    //cm.instance -> es el id del modulo
    $now         = microtime(true);
    $actividades = [];

    foreach ($cursos as $key => $value) {
        $suspendeds = $DB->get_records_sql("SELECT
                            user2.id AS id
                            FROM mdl_course AS course
                            JOIN mdl_enrol AS en ON en.courseid = course.id
                            JOIN mdl_user_enrolments AS ue ON ue.enrolid = en.id
                            JOIN mdl_user AS user2 ON ue.userid = user2.id
                            WHERE ue.status = 1
                            and course.id = " . $value->course_id);

        $cursos[$key]->students = $cursos[$key]->students - count($suspendeds);

        $type_activity = "SELECT cm.instance, cm.module
                             from {course_modules} as cm
                             INNER JOIN {course_sections} as cs ON cm.section = cs.id
                             where cs.course = $value->course_id AND cs.section = $section_course AND (cm.module = 7 or cm.module = 1)";
        //solo debe haber una encuesta o una tarea por semana
        $tipo_actividad = $DB->get_records_sql($type_activity);

        if ([] == $tipo_actividad) {
            $tipo_actividad = null;
        } else {
            $tipo_actividad = array_values($tipo_actividad)[0];
        }

        $activ = null;
        if (is_object($tipo_actividad)) {
            $activ = $tipo_actividad->module;
        }
        if (null == $activ) {
            continue;
        }
        //FLUJO DE FEEDBACK
        if (7 == $activ) {
            $sql_feedback = "SELECT fb.id,  fb.name, fb.course, c.fullname as curname, cs.section as semana , cm.module, fb.timeclose
                                from {feedback} as fb
                                join {course_modules} as cm ON fb.id = cm.instance
                                join {course_sections} as cs ON cm.section = cs.id
                                join {course} as c ON cm.course = c.id
                                where fb.course= $value->course_id AND c.id= $value->course_id AND cs.section = $section_course";

            $actividad = $DB->get_records_sql($sql_feedback);
            if ([] != $actividad) {
                foreach ($actividad as $llave => $valor) {
                    array_push($actividades, $valor);
                }
            }
            //FIN - FLUJO DE FEEDBACK
            //FLUJO DE TAREA
        } else if (1 == $activ) {
            $sql_assign = "SELECT CONCAT(ass.id,cs.id,cm.id,c.id) as fullid, ass.id,  ass.name, ass.course, c.fullname as curname, cs.section as semana , cm.module
                                from {assign} as ass
                                join {course_modules} as cm ON ass.id = cm.instance
                                join {course_sections} as cs ON cm.section = cs.id
                                join {course} as c ON cm.course = c.id
                                where ass.course= $value->course_id  AND cs.section = $section_course";

            $actividad = $DB->get_records_sql($sql_assign);

            if ([] != $actividad) {
                foreach ($actividad as $llave => $valor) {
                    array_push($actividades, $valor);
                }
            }
            //FIN - FLUJO DE TAREA
        }
    }

    /**
    calcular puntaje
     */
    $all_data = [];
    foreach ($actividades as $key => $value) {
        $cur_name = $value->curname;

        if (7 == $value->module) {
            $puntaje                = 0;
            $encuestas_cantidad_sql = "SELECT * FROM {feedback_completed} WHERE feedback = " . $value->id . " AND timemodified < " . $value->timeclose;
            $encuestas_cantidad     = $DB->get_records_sql($encuestas_cantidad_sql);

            $cContext = context_course::instance($cursos[$value->course]->course_id);

            foreach ($encuestas_cantidad as $llave => $valor) {
                $tm = array_values(get_user_roles($cContext, $valor->userid));
                if (array_key_exists($valor->userid, $suspendeds)) {
                    unset($encuestas_cantidad[$llave]);
                    continue;
                }
                $fll = true;
                foreach ($tm as $val) {
                    foreach ($role_config as $yave) {
                        if ($val->roleid == $yave) {
                            $fll = false;
                        }
                    }
                }

                if ($fll) {
                    unset($encuestas_cantidad[$llave]);
                }
            }

            if (10 == $section_course) {
                //$puntaje = (count($encuestas_cantidad) / 5) * 100;
                $puntaje = count($encuestas_cantidad);
            } else if (12 == $section_course) {
                //$puntaje = (count($encuestas_cantidad) / 1) * 100;
                $puntaje = $encuestas_cantidad ? 5 : 0;
            } else {
                //$puntaje = (count($encuestas_cantidad) / $cursos[$value->course]->students) * 100;
                $puntaje = round((count($encuestas_cantidad) * 5) / $cursos[$value->course]->students,2);
                /*
                echo "<pre>";
                print_r($puntaje);
                echo "</pre>";
                */
            }

            $fb_name = $value->name;

            $datos_feedback                = new stdClass();
            $datos_feedback->activity_name = $fb_name;
            $datos_feedback->section_name  = $section_course;
            $datos_feedback->tipo          = 'Encuesta';
            $datos_feedback->archivos      = count($encuestas_cantidad);
            $datos_feedback->puntaje       = $puntaje;

            array_push($all_data, $datos_feedback);
        } elseif (1 == $value->module) {
            $puntaje = 100;

            $sql_tarea_verification = "SELECT subass.id,subass.status,subass.userid,subass.timecreated, ass.id as 'idass', ass.name, ass.duedate
                                         from {assign} as ass
                                         join {assign_submission} as subass ON ass.id = subass.assignment
                                         where ass.id = $value->id AND status = 'submitted'";
            $tarea_verification = $DB->get_records_sql($sql_tarea_verification);

            $tarea_verification = array_values($tarea_verification);

            $valorMaximo = 0;
            foreach ($tarea_verification as $key => $valueee) {
                if ($valueee->timecreated < $valorMaximo) {
                    $tmp                      = $tarea_verification[0];
                    $tarea_verification[0]    = $tarea_verification[$key];
                    $tarea_verification[$key] = $tmp;
                }
                $valorMaximo = $valueee->timecreated;
            }
            $flag = true;
            if ([] == $tarea_verification) {
                $flag                   = false;
                $sql_tarea_verification = "SELECT ass.id as 'idass', ass.name, ass.duedate
                                         from {assign} as ass
                                         where ass.id = $value->id";
                $tarea_verification = $DB->get_records_sql($sql_tarea_verification);

                $tarea_verification = array_values($tarea_verification);
            }

            $archivos = 0;
            foreach ($tarea_verification as $key => $vaa) {
                if (isset($vaa->id) && '' != $vaa->id) {
                    $ddd = $DB->get_records_sql("SELECT *  FROM {files} WHERE itemid = " . $vaa->id . " AND userid = " . $vaa->userid . " and filesize > 0");
                    $archivos += count($ddd);
                }

                $cm = get_coursemodule_from_instance('assign', $tarea_verification[$key]->idass);
                if (!isset($vaa->userid)) {
                    continue;
                }
                $loog = [];
                //$loog = array_values( $DB->get_records('logstore_standard_log',  array('objecttable' => 'assign_submission',
                //                                                      'userid' => $tarea_verification[$key]->userid,
                //                                                       'action' => 'uploaded',
                //                                                       'contextinstanceid' => $cm->id)) );
                if ([] == $loog) {
                    continue;
                }
                $tarea_verification[$key]->timecreated = (null != $loog[count($loog) - 1]->timecreated) ? $loog[count($loog) - 1]->timecreated : $tarea_verification[$key]->timecreated;
            }

            usort($tarea_verification, function ($a, $b) {
                return strcmp($a->timecreated, $b->timecreated);
            });

            //$archivos = count($tarea_verification);
            while (count($tarea_verification) > 1) {
                array_pop($tarea_verification);
            }

            if ($flag) {
                $cm   = get_coursemodule_from_instance('assign', $tarea_verification[0]->idass);
                $loog = [];
                //$loog = array_values( $DB->get_records('logstore_standard_log',  array('objecttable' => 'assign_submission',
                //                                                        'userid' => $tarea_verification[0]->userid,
                //                                                         'action' => 'uploaded',
                //                                                         'contextinstanceid' => $cm->id)) );
                $tarea_verification[0]->newtimecreated = date('d/m/y h:i:s A', $tarea_verification[0]->timecreated);
                $tarea_verification[0]->newduedate     = date('d/m/y h:i:s A', $tarea_verification[0]->duedate);
            } else {
                $loog                 = [];
                $loog[0]              = new stdClass();
                $loog[0]->timecreated = null;
            }

            if ([] == $loog) {
                if (!isset($tarea_verification[0]->timecreated) || null == $tarea_verification[0]->timecreated) {
                    $tarea_verification[0]->timecreated = $tarea_verification[0]->duedate + 900000000;
                }
            } else {
                $tarea_verification[0]->timecreated = (null != $loog[count($loog) - 1]->timecreated) ? $loog[count($loog) - 1]->timecreated : ($tarea_verification[0]->duedate + 900000000);
            }

            foreach ($tarea_verification as $key => $vaa) {
                $dias = ($vaa->duedate - $vaa->timecreated) / 86400;

                $dias = $dias * -1;

                $dias = number_format($dias, 2);

                if ($dias <= 0) {
                    $puntaje = 5;
                }

                //$archivos = 1;
                switch (ceil($dias)) {
                    case 1:
                        $puntaje = 4;
                        break;
                    case 2:
                        $puntaje = 3;
                        break;
                    case 3:
                        $puntaje = 2;
                        break;
                    case 4:
                        $puntaje = 1;
                        break;
                }
                if (ceil($dias) > 4) {
                    $puntaje = 0;
                }

                if (!$flag) {
                    $puntaje  = 0;
                    $archivos = 0;
                }

                $grade_item = $DB->get_record('grade_items', ['iteminstance' => $cm->instance, 'itemmodule' => 'assign', 'courseid' => $cm->course ]);
                
                $grade_grade = $DB->get_record('grade_grades',['itemid' => $grade_item->id, 'userid' => $vaa->userid]);
                
                if($grade_grade != null){
                    $grade = round($grade_grade->finalgrade);
                    if($grade <= 5){
                        $puntaje == $grade;
                    }
                }

                $ass_name = $vaa->name;

                $datos_tarea                = new stdClass();
                $datos_tarea->activity_name = $ass_name;
                $datos_tarea->section_name  = $section_course;
                $datos_tarea->tipo          = 'Tarea';
                $datos_tarea->puntaje       = $puntaje;
                $datos_tarea->archivos      = $archivos;

                array_push($all_data, $datos_tarea);
            }
        }
    }
    /*
    if (152 == $courseid and 13 == $section_course) {
    echo "<pre>";
    print_r($all_data);
    echo "</pre>";
    die();
    }
    echo "<pre>";
    print_r($all_data);
    echo "</pre>";
    die();
     */
    return $all_data;
}

function db_data($value)
{
    global $DB, $CFG;
    require_once $CFG->libdir . '/adminlib.php';
    require_once $CFG->libdir . '/modinfolib.php';
    require_once $CFG->libdir . '/formslib.php';

    $categoryid     = $courseid;
    $section_course = $section_course;
    $role_config    = [$role_config];

    $archivos = 0;

    $sql_tarea_verification = "SELECT subass.id,subass.status,subass.userid,subass.timecreated, ass.id as 'idass', ass.name, ass.duedate
                                    from {assign} as ass
                                    join {assign_submission} as subass ON ass.id = subass.assignment
                                    where ass.id = $value AND status = 'submitted'";
    $tarea_verification = $DB->get_records_sql($sql_tarea_verification);

    $tarea_verification = array_values($tarea_verification);

    $valorMaximo = 0;
    foreach ($tarea_verification as $key => $valueee) {
        if ($valueee->timecreated < $valorMaximo) {
            $tmp                      = $tarea_verification[0];
            $tarea_verification[0]    = $tarea_verification[$key];
            $tarea_verification[$key] = $tmp;
        }
        $valorMaximo = $valueee->timecreated;
    }
    $flag = true;
    if ([] == $tarea_verification) {
        $flag                   = false;
        $sql_tarea_verification = "SELECT ass.id as 'idass', ass.name, ass.duedate
                                    from {assign} as ass
                                    where ass.id = $value";
        $tarea_verification = $DB->get_records_sql($sql_tarea_verification);

        $tarea_verification = array_values($tarea_verification);
    }

    foreach ($tarea_verification as $key => $vaa) {
        if (isset($vaa->id) && '' != $vaa->id) {
            $ddd = $DB->get_records_sql("SELECT *  FROM {files} WHERE itemid = " . $vaa->id . " AND userid = " . $vaa->userid . " and filesize > 0");
            $archivos += count($ddd);
        }
    }

    usort($tarea_verification, function ($a, $b) {
        return strcmp($a->timecreated, $b->timecreated);
    });

    return $archivos;

}

function suspended_data($courses, $categoryid)
{
    global $DB;

    $glob_roles = [];
    $time_start_ts = time();
    foreach ($courses as $key => $value) {
        $tmp_data = explode('-', $value->fullname);

        unset($value->fullname);
        $split_name = explode(' DE ', $tmp_data[0]);
        $value->departamento = $split_name[count($split_name) - 1];
        $value->name         = $tmp_data[0];
        $value->equipo       = isset($tmp_data[3]) ? $tmp_data[3] : '';
        $value->categoria    = isset($tmp_data[4]) ? $tmp_data[4] : '';

        $status_enrol = '';
        if (1 == $value->visible) {
            $status_enrol = " AND ue.status = 1 ";
        }

        $sql_roles = "SELECT concat(u.id,r.id,ra.id) as ignoreid,u.id as userid, u.firstname ,u.lastname, u.email, u.address,u.username as dni, r.id as roleid ,r.shortname as rolename, u.phone2, u.city as departamento, c.id as courseid
                  FROM mdl_user u
                  JOIN mdl_user_enrolments ue ON ue.userid = u.id
                  JOIN mdl_enrol e ON e.id = ue.enrolid
                  JOIN mdl_role_assignments ra ON ra.userid = u.id
                  JOIN mdl_context ct ON ct.id = ra.contextid AND ct.contextlevel = 50
                  JOIN mdl_course c ON c.id = ct.instanceid AND e.courseid = c.id
                  JOIN mdl_role r ON r.id = ra.roleid
                  WHERE  c.id = " . $value->id . "
                  AND e.status = 0 AND u.suspended = 0 AND u.deleted = 0
                   " . $status_enrol . " ORDER BY u.lastname ASC";

        $sql = $DB->get_records_sql($sql_roles);

        $value->users = [];

        $sql_tot_splash = "SELECT concat(u.id,r.id,ra.id) as ignoreid,u.id as userid, u.firstname ,u.lastname,u.username as dni, u.email, u.address, r.id as roleid ,r.shortname as rolename, u.city as departamento, c.id as courseid
                  FROM mdl_user u
                  JOIN mdl_user_enrolments ue ON ue.userid = u.id
                  JOIN mdl_enrol e ON e.id = ue.enrolid
                  JOIN mdl_role_assignments ra ON ra.userid = u.id
                  JOIN mdl_context ct ON ct.id = ra.contextid AND ct.contextlevel = 50
                  JOIN mdl_course c ON c.id = ct.instanceid AND e.courseid = c.id
                  JOIN mdl_role r ON r.id = ra.roleid AND r.id = " . SP_STUDENT . "
                  WHERE  c.id = " . $value->id . "
                  AND e.status = 0 AND u.suspended = 0 AND u.deleted = 0
                   " . $status_enrol . " ORDER BY u.lastname ASC";
        $tot_splash = $DB->get_records_sql($sql_tot_splash);

        $prof = 0;

        $roles = [];

        foreach ($sql as $ke => $va) {
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
                if ($ke == $repe) {
                    $irse = true;
                }
            }
            if ($irse) {
                continue;
            }

            $name_roles = $roles;

            if (SP_TEACHER_NO_EDITING != $va->roleid) {
                $name_roles[$va->roleid]          = $name_roles[$va->roleid];
                $name_roles[$va->roleid]['valeu'] = '1';
                //unset($name_roles[$va->roleid]);
            }

            if (SP_TEACHER_NO_EDITING == $va->roleid) {
                $prof                             = $va->userid;
                $va->spl                          = count($tot_splash);
                $name_roles[$va->roleid]          = $name_roles[SP_TEACHER_NO_EDITING];
                $name_roles[$va->roleid]['valeu'] = '1';
                //unset($name_roles[SP_TEACHER_NO_EDITING]);
                $tmp_key_c = $ke;
            } else if ($prof != $va->userid) {
                $va->spl = '';
            }

            foreach ($sql as $ksql => $vsql) {
                if ($va->userid == $vsql->userid && $va->ignoreid != $vsql->ignoreid) {
                    if (SP_TEACHER_NO_EDITING != $vsql->roleid) {
                        $name_roles[$vsql->roleid]          = $name_roles[$vsql->roleid];
                        $name_roles[$vsql->roleid]['valeu'] = '1';
                        //unset($name_roles[$vsql->roleid]);
                    }

                    if (SP_TEACHER_NO_EDITING == $vsql->roleid) {
                        $prof                               = $vsql->userid;
                        $sql[$ke]->spl                      = count($tot_splash);
                        $va->spl                            = count($tot_splash);
                        $name_roles[$vsql->roleid]          = $name_roles[SP_TEACHER_NO_EDITING];
                        $name_roles[$vsql->roleid]['valeu'] = '1';
                        //unset($name_roles[SP_TEACHER_NO_EDITING]);
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
            if (is_object($totLogins)) {
                $va->totlogins = $totLogins->total;
            } else {
                $va->totlogins = 0;
            }

            $data    = [];
            $puntaje = 0;
            $nsemana = $DB->get_record('config', ['name' => 'reportpointscantsem'])->value;
            for ($i = 2; $i < $nsemana; $i++) {
                $exist_data = $DB->get_record('local_wsplashdata', ['courseid' => $value->id, 'semana' => $i]);
                if (is_object($exist_data)) {
                    $tmp_data = json_decode($exist_data->data);
                    if ('Tarea' == $tmp_data[0]->tipo) {
                        $tmp_data[0]->archivos = db_data($value->id, $i, SP_STUDENT);
                    }
                } else {
                    $tmp_data = ws_data($value->id, $i, SP_STUDENT);
                }

                if ([] == $tmp_data || !isset($tmp_data[0]->section_name)) {
                    //unset($data['Semana ' . $i]);
                    continue;
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
                    //$v->archivos = ($v->puntaje == '') ? '' : '1';
                    //$v->archivos = ($v->puntaje == 0) ? '1' : $v->archivos;
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
                //continue;
                if ([] == $data) {
                    continue;
                }

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
                $va->totlogins    = '';
                $va->puntajetotal = '';
            }
        }
        $sql_wi = [];
        if ([] != $sql && '' != $tmp_key_c) {
            $sql_wi['temporal'] = $sql[$tmp_key_c];
            unset($sql[$tmp_key_c]);
            array_unshift($sql, $sql_wi['temporal']);
        }
        $value->users[] = array_values($sql);
        //if (intval($key) > 10) {
        //    echo var_dump(time() - $time_start_ts);
        //    exit;
        //}
    }
    return [$courses, $glob_roles];
}

function firstpage($objPHPExcel, $glob_roles, $columnos_excel, $output, $max, $nombreHoja, $tmp_roles2,$suspend = false)
{
    require_once '../classes/PHPExcel.php';

    $role_s_default = '';
    $role_c_default = '';
    $header_excel   = 0;
    $objPHPExcel
        ->setCellValue($columnos_excel[$header_excel++] . '2', '#')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Auditor')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Control')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Departamento')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Nombre del Colegio')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Aula')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Equipo')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Categoría')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Apellido')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Nombre')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'DNI')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Celular')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Email')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Fnac')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'SPL')
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'Rol');
    $excel_roles = [];
    foreach ($glob_roles as $key => $value) {
        if ('s' == $key) {
            $role_s_default = $columnos_excel[$header_excel];
        }
        if ('c' == $key) {
            $role_c_default = $columnos_excel[$header_excel];
        }
        foreach ($tmp_roles2 as $ll => $vaa) {
            if ($ll == $key) {
                $tmp_roles2[$columnos_excel[$header_excel]] = '';
                unset($tmp_roles2[$ll]);
            }
        }
        $excel_roles[$key] = $columnos_excel[$header_excel];
        $objPHPExcel
            ->setCellValue($columnos_excel[$header_excel++] . '2', strtoupper($key));
    }
    $objPHPExcel
        ->setCellValue($columnos_excel[$header_excel++] . '2', 'LOG');
    for ($i = 2; $i <= $max; $i++) {
        //echo $columnos_excel[$header_excel] . '2 --- Semana ' . $i . '<br>';
        $first_column  = $header_excel++;
        $second_column = $header_excel++;
        $objPHPExcel
            ->setCellValue($columnos_excel[$first_column] . '1', 'SEM ' . $i);
        //$header_excel++;
        $objPHPExcel
            ->mergeCells($columnos_excel[$first_column] . '1:' . $columnos_excel[$second_column] . '1');
        $objPHPExcel
            ->setCellValue($columnos_excel[$first_column] . '2', 'PUNT')
            ->setCellValue($columnos_excel[$second_column] . '2', (10 == $i || 12 == $i || 16 == $i) ? 'RPTA' : 'ARCH');
    }
    $first_column  = $header_excel++;
    $second_column = $header_excel++;

    $objPHPExcel
        ->setCellValue($columnos_excel[$first_column] . '1', 'SEM 16');
    $objPHPExcel
        ->mergeCells($columnos_excel[$first_column] . '1:' . $columnos_excel[$second_column] . '1');

    $objPHPExcel
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
        $tmp_index = $index_con;
        $index_con++;
        $tmpDep      = '';
        $repeat_role = [];
        foreach ($value->users[0] as $val) {
            $column_data = 0;

            if (strlen($val->totlogins) >= 1 || $suspend) {
                //$tmpDep = $val->departamento;

                $objPHPExcel
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $tmp_index)
                    ->setCellValue($columnos_excel[$column_data++] . $actual, explode("-", $value->idnumber)[2] ?? '')
                    ->setCellValue($columnos_excel[$column_data++] . $actual, '')
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $value->departamento)
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $value->name)
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $value->aula)
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $value->equipo)
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $value->categoria);
            } else {
                $column_data++;
                $column_data++;
                $objPHPExcel
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $control_acronim_reporter[$value->aula] ?? '');
                $column_data++;
                $column_data++;
                $column_data++;
                $column_data++;
                $column_data++;
            }

            $objPHPExcel
                ->setCellValue($columnos_excel[$column_data++] . $actual, $val->rolfulname != "Consejero" ? mb_strtoupper($val->lastname, 'utf-8') : ucwords($val->lastname))
                ->setCellValue($columnos_excel[$column_data++] . $actual, $val->rolfulname != "Consejero" ? mb_strtoupper($val->firstname, 'utf-8') : ucwords($val->firstname))
                ->setCellValue($columnos_excel[$column_data] . $actual, (string) $val->dni);
            $objPHPExcel->setCellValueExplicit(
                $columnos_excel[$column_data++] . $actual,
                $val->dni,
                PHPExcel_Cell_DataType::TYPE_STRING
            );
            $objPHPExcel
                ->setCellValue($columnos_excel[$column_data++] . $actual, $val->phone2)
                ->setCellValue($columnos_excel[$column_data++] . $actual, $val->email)
                ->setCellValue($columnos_excel[$column_data++] . $actual, $val->address)
                ->setCellValue($columnos_excel[$column_data++] . $actual, $val->spl)
                ->setCellValue($columnos_excel[$column_data++] . $actual, $val->rolfulname == "Consejero" ? $val->rolfulname : "Splasher");
            $tmp_index = '';

            foreach ($val->roles as $rolee) {
                $objPHPExcel
                    ->setCellValue($excel_roles[$rolee['name']] . $actual, $rolee['valeu']);
                if (1 == $rolee['valeu'] && $excel_roles[$rolee['name']] != $role_s_default) {
                    if (isset($repeat_role[$excel_roles[$rolee['name']]])) {
                        cellColor($excel_roles[$rolee['name']] . $actual, 'F28A8C', $objPHPExcel);
                        cellColor($repeat_role[$excel_roles[$rolee['name']]], 'F28A8C', $objPHPExcel);
                    } else {
                        $repeat_role[$excel_roles[$rolee['name']]] = $excel_roles[$rolee['name']] . $actual;
                    }
                }

                if ($excel_roles[$rolee['name']] != $role_c_default && 1 == $rolee['valeu'] && '' != $val->spl) {
                    cellColor($excel_roles[$rolee['name']] . $actual, 'F28A8C', $objPHPExcel);
                }

                if (isset($tmp_roles2[$excel_roles[$rolee['name']]]) && 1 == $rolee['valeu']) {
                    cellColor($excel_roles[$rolee['name']] . $actual, 'F28A8C', $objPHPExcel);
                }
            }
            $column_data += count($excel_roles);

            $objPHPExcel
                ->setCellValue($columnos_excel[$column_data++] . $actual, $val->totlogins);

            foreach ($val->data as $puntajee) {
                $objPHPExcel
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $puntajee[0]->puntaje);
                $objPHPExcel
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $puntajee[0]->archivos);
            }

            if (!empty($val->totlogins)) {
                global $DB;
                $custom_fields = [];

                $sql_custom_fields = "
                    SELECT muif.shortname, muid.data
                    FROM mdl_role_assignments AS ra 
                    JOIN mdl_user_enrolments AS ue ON ra.userid = ue.userid 
                    JOIN mdl_role AS r ON ra.roleid = r.id 
                    JOIN mdl_context AS c ON c.id = ra.contextid 
                    JOIN mdl_enrol AS e ON e.courseid = c.instanceid AND ue.enrolid = e.id 
                    JOIN mdl_user_info_data muid ON muid.userid = ra.userid
                    JOIN mdl_user_info_field muif ON muif.id = muid.fieldid  
                    WHERE e.courseid = $val->courseid AND r.shortname = 'p';
                ";

                $sql_custom_fields = "SELECT gi.idnumber as 'shortname', ROUND(gg.finalgrade) as 'data' FROM {grade_grades} gg
                                      INNER JOIN {grade_items} gi ON  gi.id = gg.itemid 
                                      WHERE (gi.idnumber = 'audit1' OR  gi.idnumber = 'audit2') and gg.userid = " . $val->userid;
                $res_custom_fields = $DB->get_records_sql($sql_custom_fields);
                if($res_custom_fields == []){
                    $res_custom_fields = ['audit1' => '', 'audit2' => ''];
                }
                
                foreach ($res_custom_fields as $rcf) {
                    $custom_fields[$rcf->shortname] = $rcf->data; 
                };

                $objPHPExcel
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $custom_fields["audit1"] ?? "");
                $objPHPExcel
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $custom_fields["audit2"] ?? "");
                $objPHPExcel
                    ->setCellValue($columnos_excel[$column_data++] . $actual, $val->puntajetotal + intval($custom_fields["audit1"] ?? "") + intval($custom_fields["audit2"] ?? ""));
            } else {
                $column_data++;
                $column_data++;
                $column_data++;
            }
            $actual++;
        }
    }

    $objPHPExcel->freezePaneByColumnAndRow(0, 3);

    $objPHPExcel->setTitle($nombreHoja);
}
