<?php
function suspended_data($courses, $categoryid) {
    global $DB;

    $glob_roles = [];
    foreach ($courses as $key => $value) {
        $tmp_data = explode('-', $value->fullname);

        $split_name = explode(' DE ', $tmp_data[0]);
        $value->departamento = $split_name[count($split_name) - 1];
        $value->name         = $tmp_data[0];
        $value->equipo       = isset($tmp_data[3]) ? $tmp_data[3] : '';
        $value->categoria    = isset($tmp_data[4]) ? $tmp_data[4] : '';

        $status_enrol = (1 == $value->visible) ? " AND ue.status = 1 " : "";

        $sql_roles = "SELECT concat(u.id,r.id,ra.id) as ignoreid,u.id as userid, u.firstname ,u.lastname, u.email, u.address,u.username as dni, r.id as roleid ,r.shortname as rolename, u.phone2, u.city as departamento
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

        $prof = 0;
        $roles = [];

        foreach ($sql as $ke => $va) {
            if (SP_TEACHER_NO_EDITING == $va->roleid) {
                $prof = $va->userid;
            }
            $roles[$va->roleid]        = ['name' => $va->rolename, 'valeu' => ''];
            $glob_roles[$va->rolename] = '';
        }

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
            if ($irse) continue;

            $name_roles = $roles;
            if (SP_TEACHER_NO_EDITING != $va->roleid) {
                $name_roles[$va->roleid]['valeu'] = '1';
            }
            if (SP_TEACHER_NO_EDITING == $va->roleid) {
                $prof                             = $va->userid;
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
                        $prof                               = $vsql->userid;
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

            $data    = [];
            $puntaje = 0;
            $nsemana = 17;
            for ($i = 2; $i < $nsemana; $i++) {
                $tmp_data = ws_data($value->id, $i, SP_STUDENT);

                if ([] == $tmp_data || !isset($tmp_data[0]->section_name)) continue;

                $data['Semana ' . $tmp_data[0]->section_name] = $tmp_data;

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
        }
        $sql_wi = [];
        if ([] != $sql && '' != $tmp_key_c) {
            $sql_wi['temporal'] = $sql[$tmp_key_c];
            unset($sql[$tmp_key_c]);
            array_unshift($sql, $sql_wi['temporal']);
        }
        $value->users[] = array_values($sql);
    }
    return [$courses, $glob_roles];
}