<?php
function db_data($value) {
    global $DB, $CFG;
    require_once $CFG->libdir . '/adminlib.php';
    require_once $CFG->libdir . '/modinfolib.php';
    require_once $CFG->libdir . '/formslib.php';

    $archivos = 0;
    $sql_tarea_verification = "SELECT subass.id,subass.status,subass.userid,subass.timecreated, ass.id as 'idass', ass.name, ass.duedate
                                    from {assign} as ass
                                    join {assign_submission} as subass ON ass.id = subass.assignment
                                    where ass.id = $value AND status = 'submitted'";
    $tarea_verification = $DB->get_records_sql($sql_tarea_verification);

    $tarea_verification = array_values($tarea_verification);

    foreach ($tarea_verification as $key => $vaa) {
        if (isset($vaa->id) && '' != $vaa->id) {
            $ddd = $DB->get_records_sql("SELECT *  FROM {files} WHERE itemid = " . $vaa->id . " AND userid = " . $vaa->userid . " and filesize > 0");
            $archivos += count($ddd);
        }
    }
    return $archivos;
}