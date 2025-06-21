<?php
function ws_data($courseid, $section_course, $role_config) {
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
        }
    }

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
                $puntaje = count($encuestas_cantidad);
            } else if (12 == $section_course) {
                $puntaje = $encuestas_cantidad ? 5 : 0;
            } else {
                $puntaje = round((count($encuestas_cantidad) * 5) / $cursos[$value->course]->students,2);
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
            }

            usort($tarea_verification, function ($a, $b) {
                return strcmp($a->timecreated, $b->timecreated);
            });

            if ($flag) {
                $tarea_verification[0]->newtimecreated = date('d/m/y h:i:s A', $tarea_verification[0]->timecreated);
                $tarea_verification[0]->newduedate     = date('d/m/y h:i:s A', $tarea_verification[0]->duedate);
            }

            if ([] == $tarea_verification) {
                if (!isset($tarea_verification[0]->timecreated) || null == $tarea_verification[0]->timecreated) {
                    $tarea_verification[0]->timecreated = $tarea_verification[0]->duedate + 900000000;
                }
            } else {
                $tarea_verification[0]->timecreated = (null != $tarea_verification[count($tarea_verification) - 1]->timecreated) ? $tarea_verification[count($tarea_verification) - 1]->timecreated : ($tarea_verification[0]->duedate + 900000000);
            }

            foreach ($tarea_verification as $key => $vaa) {
                $dias = ($vaa->duedate - $vaa->timecreated) / 86400;
                $dias = $dias * -1;
                $dias = number_format($dias, 2);

                if ($dias <= 0) {
                    $puntaje = 5;
                }
                switch (ceil($dias)) {
                    case 1: $puntaje = 4; break;
                    case 2: $puntaje = 3; break;
                    case 3: $puntaje = 2; break;
                    case 4: $puntaje = 1; break;
                }
                if (ceil($dias) > 4) $puntaje = 0;
                if (!$flag) {
                    $puntaje  = 0;
                    $archivos = 0;
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
    return $all_data;
}