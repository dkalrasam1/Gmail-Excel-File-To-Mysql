<?php
 
 
/**
 * 
 *	Gmail attachment extractor.
 *
 *	Downloads attachments from Gmail and saves it to a file.
 *	Uses PHP IMAP extension, so make sure it is enabled in your php.ini,
 *	extension=php_imap.dll
 *
 */
 
 
set_time_limit(3000); 
 
 
/* connect to gmail with your credentials */
$hostname = '{imap.gmail.com:993/imap/ssl/novalidate-cert}INBOX';
//$hostname = '{imap.gmail.com:993/imap/ssl}INBOX';
$username = 'xxxxxxxxxxx'; # e.g somebody@gmail.com
$password = 'xxxxxxxxxxx';
 
 
/* try to connect */
$inbox = imap_open($hostname,$username,$password) or die('Cannot connect to Gmail: ' . imap_last_error());
 
 
/* get all new emails. If set to 'ALL' instead 
 * of 'NEW' retrieves all the emails, but can be 
 * resource intensive, so the following variable, 
 * $max_emails, puts the limit on the number of emails downloaded.
 * 
 */
$emails = imap_search($inbox,'FROM "xxxxxxxxxxx" UNSEEN');

 
/* useful only if the above search is set to 'ALL' */
$max_emails = 16;
 
 //echo "<pre>";
/* if any emails found, iterate through each email */
$allFiles = [];
if($emails) {
 
    $count = 1;
 
    /* put the newest emails on top */
    rsort($emails);

    
    /* for every email... */
    foreach($emails as $email_number) 
    {
 
        /* get information specific to this email */
        $overview = imap_fetch_overview($inbox,$email_number,0);
        //print_r($overview);
        /* get mail message */
        $message = imap_fetchbody($inbox,$email_number,2);
 
        /* get mail structure */
        $structure = imap_fetchstructure($inbox, $email_number);
 
        $attachments = array();
        //print_r($structure);
        /* if any attachments found... */
        if(isset($structure->parts) && count($structure->parts)) 
        {
            for($i = 0; $i < count($structure->parts); $i++) 
            {
                $attachments[$i] = array(
                    'is_attachment' => false,
                    'filename' => '',
                    'name' => '',
                    'attachment' => ''
                );
 
                if($structure->parts[$i]->ifdparameters && isset($structure->parts[$i]->dparameters)) 
                {
                    foreach($structure->parts[$i]->dparameters as $object) 
                    {
                        if(strtolower($object->attribute) == 'filename') 
                        {
                            $attachments[$i]['is_attachment'] = true;
                            $attachments[$i]['filename'] = $object->value;
                        }
                    }
                }
 
                // if($structure->parts[$i]->ifparameters) 
                // {
                //     foreach($structure->parts[$i]->parameters as $object) 
                //     {
                //         if(strtolower($object->attribute) == 'name') 
                //         {
                //             $attachments[$i]['is_attachment'] = true;
                //             $attachments[$i]['name'] = $object->value;
                //         }
                //     }
                // }
 
                if($attachments[$i]['is_attachment']) 
                {
                    $attachments[$i]['attachment'] = imap_fetchbody($inbox, $email_number, $i+1);
                    
                    /* 4 = QUOTED-PRINTABLE encoding */
                    if($structure->parts[$i]->encoding == 3) 
                    { 
                        $attachments[$i]['attachment'] = base64_decode($attachments[$i]['attachment']);
                    }
                    /* 3 = BASE64 encoding */
                    elseif($structure->parts[$i]->encoding == 4) 
                    { 
                        $attachments[$i]['attachment'] = quoted_printable_decode($attachments[$i]['attachment']);
                    }

                    if(empty($filename)) $filename = time() . ".dat";
 
                    /* prefix the email number to the filename in case two emails
                    * have the attachment with the same file name.
                    */
                    
                    $filename = $attachments[$i]['name'];
                    if(empty($filename)) $filename = $attachments[$i]['filename'];
                    $fp = fopen($email_number . "-" . $filename, "w+");
                    fwrite($fp, $attachments[$i]['attachment']);
                    array_push($allFiles,$email_number . "-" . $filename);
                    fclose($fp);
                
                }
            }
        }
 
        /* iterate through each attachment and save it */
        //print_r($attachments);
        // foreach($attachments as $attachment)
        // {
        //     if($attachment['is_attachment'] == 1)
        //     {
        //         $filename = $attachment['name'];
        //         if(empty($filename)) $filename = $attachment['filename'];
 
        //         if(empty($filename)) $filename = time() . ".dat";
 
        //         /* prefix the email number to the filename in case two emails
        //          * have the attachment with the same file name.
        //          */
        //         $fp = fopen($email_number . "-" . $filename, "w+");
        //         fwrite($fp, $attachment['attachment']);
        //         fclose($fp);
        //     }
 
        // }
 
        if($count++ >= $max_emails) break;
    }
 
} 
 
/* close the connection */
imap_close($inbox);
if(count($allFiles) > 0){
    require_once __DIR__.'/simple-xlsx/src/SimpleXLSX.php';

    $servername = "localhost";
    $username = "root";
    $password = "password";
    $dbname = "db";

    // Create connection
    $conn = new mysqli($servername, $username, $password, $dbname);
    // Check connection
    if ($conn->connect_error) {
        error_log($conn->connect_error, 3, "/var/www/html/gmailFetch/cron.log");
    die("Connection failed: " . $conn->connect_error);
    }
    $error = '';
    /*****************************************************************
     * 
     * find last id in table starts
     * 
     ******************************************************************/

    $sql = "select id from calls order by id desc limit 1";
    $start_id_result = $conn->query($sql);
    $start_id_data = $start_id_result->fetch_assoc();

    if( empty( $start_id_data ) )
        $start_id = 1;
    else
        $start_id = $start_id_data['id']+1;

    /*****************************************************************
     * 
     * find last id in table ends
     * 
     ******************************************************************/



    /*****************************************************************
     * 
     * Reading data from xlsx and inserting users starts
     * 
     ******************************************************************/
    $insert_query = 'INSERT IGNORE INTO `calls`(`session_id`, `from_name`, `from_no`, `to_name`, `to_no`, `result`, `call_length`, `handle_time`, `call_start_time`, `call_direction`, `queue`) VALUES';
    if ( $xlsx = SimpleXLSX::parse($allFiles[0]) ) {
        $j = 0;
        $total_records = count($xlsx->rows(1))-1;
        if( $total_records >= 1){
            foreach($xlsx->rows(1) as $row){
                if($j != 0){
                    $insert_query.="('$row[0]', '$row[1]', '$row[2]', '$row[3]', '$row[4]', '$row[5]', '$row[6]', '$row[7]', '$row[8]', '$row[9]', '$row[10]'),";
                }
                $j++;
            }
            $insert_query = rtrim($insert_query, ", ");
            if ($conn->query($insert_query) === TRUE) {
            // $last_id = $conn->insert_id;
            } else {
                error_log($conn->connect_error, 3, "/var/www/html/gmailFetch/cron.log");
            }
        
        }
    } else {
        echo SimpleXLSX::parseError();
    }

    /*****************************************************************
     * 
     * Reading data from xlsx and inserting users ends
     * 
     ******************************************************************/


    /*****************************************************************
     * 
     * find last inserted id in table starts
     * 
     ******************************************************************/

    $sql = "select id from calls order by id desc limit 1";
    $last_id_result = $conn->query($sql);
    $last_id_data = $last_id_result->fetch_assoc();

    if( empty( $last_id_data ) )
        $last_id = 1;
    else
        $last_id = $last_id_data['id'];

    /*****************************************************************
     * 
     * find last inserted id in table ends
     * 
     ******************************************************************/



    /*****************************************************************
     * 
     * Insert log starts
     * 
     ******************************************************************/
    $inserted_rows = ($last_id - $start_id )+1;
    $log_query = "INSERT INTO `logs`(`file_name`, `total_rows`, `inserted_rows`, `error`, `start_id`, `end_id`, `type`) VALUES('$allFiles[0]',$total_records,$inserted_rows,'$error',$start_id,$last_id, 'calls')";
    if ($conn->query($log_query) === TRUE) {
    // $last_id = $conn->insert_id;
    } else {
        error_log($conn->connect_error, 3, "/var/www/html/gmailFetch/cron.log");
    }

    /*****************************************************************
     * 
     * Insert log ends
     * 
     ******************************************************************/
}
die;
 
?>