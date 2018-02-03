<?php
// Migration script for converting What-Do-You-Know (wdk) MS Access DB articles
// and associated comments into Wordpress-importable csv.
//
// Author: Will Mooreston
// Created: 2009-??-??
//

$hostname = #####;
$database = #####;
$username = #####;
$password = #####;
$db_conn = mysql_pconnect($hostname, $username, $password) or trigger_error(mysql_error(),E_USER_ERROR);
mysql_select_db($database, $db_conn);

$knowbits_arr = get_knowbits($db_conn);

foreach ($knowbits_arr as $knowbit_id => $knowbit_id_arr) {
    $comments = get_comments($db_conn, $knowbit_id);
    $knowbits_arr[$knowbit_id]["comments"] = $comments;

    $categories = get_categories($db_conn, $knowbit_id);
    $knowbits_arr[$knowbit_id]["categories"] = $categories;
}

//var_dump($knowbits_arr);

$comment_headers = get_comment_headers($db_conn);
$headers = '"csv_post_title", '.
            '"csv_post_post", '.
            '"csv_post_categories", '.
            '"csv_post_tags", '.
            '"csv_post_author", '.
            '"csv_post_date", '.
            $comment_headers;

echo $headers."\n";
foreach ($knowbits_arr as $knowbit_id => $knowbit_id_arr) {
    echo '"'.$knowbit_id_arr["title"].'", ';
    echo '"'.$knowbit_id_arr["post"].'", ';
    echo $knowbit_id_arr["categories"];
    echo '"'.$knowbit_id_arr["tags"].'", ';
    echo '"'.$knowbit_id_arr["author"].'", ';
    echo '"'.$knowbit_id_arr["date"].'", ';
    echo $knowbit_id_arr["comments"];
    echo "\n";
}
echo "\n";


######################################################################
# GET_KNOWBITS
######################################################################
function get_knowbits($db_conn) {
$sql_all_knowbits = "
SELECT  id AS knowbit_id,
                title,
                abstract,
                description,
                keywords,
                location,
                phone,
                url,
                username,
                submitted_when
FROM knowbits"; // WHERE id in ('1156')"; //TEST article

    $rs_knowbits = mysql_query($sql_all_knowbits, $db_conn);
    $row_rs_knowbits = mysql_fetch_assoc($rs_knowbits);
    $totalRows_rs_knowbits = mysql_num_rows($rs_knowbits);

    //echo "total knowbit rows: $totalRows_rs_knowbits\n";

    $arr = array();
    do {
        $knowbit_id = $row_rs_knowbits['knowbit_id'];
        $title = add_extra_quotes($row_rs_knowbits['title']);
        $abstract = $row_rs_knowbits['abstract'];
        $description = $row_rs_knowbits['description'];
        $keywords = $row_rs_knowbits['keywords'];
        $location = $row_rs_knowbits['location'];
        $phone = $row_rs_knowbits['phone'];
        $url = fix_urls($row_rs_knowbits['url']);
        $username = $row_rs_knowbits['username'];
        $submitted_when = $row_rs_knowbits['submitted_when'];

        // Turn gathered text bits into Wordpress-ready article text
        $post = merge_into_post($abstract,$description,$location,$phone,$url);
        $post = add_extra_quotes($post);

        // Finishing touches
        $arr[$knowbit_id]["title"] = $title;
        $arr[$knowbit_id]["post"] = $post;
        $arr[$knowbit_id]["tags"] = $keywords;
        $arr[$knowbit_id]["author"] = $username;
        $arr[$knowbit_id]["date"] = $submitted_when;

    } while ($row_rs_knowbits = mysql_fetch_assoc($rs_knowbits));

    return $arr;

}


######################################################################
# FIX_URLS: adds xhtml tags (joesweb just has raw url w/o tags, adding tags in the code)
######################################################################
function fix_urls($url_string) {
    $urls = array();
    $urls = explode(",", $url_string);

    $fixed_urls = "<ul>";

    foreach ($urls as $url) {
        $url = preg_replace("/ /", "", $url); //strip spaces
        $new_url = '<li><a href="'.$url.'" target="_blank" />'.$url.'</a></li>';
        $fixed_urls .= $new_url;
    }

    $fixed_urls .= "</ul>";

    return $fixed_urls;

}


######################################################################
# GET_COMMENT_HEADERS
######################################################################
function get_comment_headers($db_conn) {
    $comment_headers = '';
$sql_max_headers = "
SELECT MAX( kbCount ) AS total
FROM (
    SELECT count( knowbit_id ) AS kbCount
    FROM comments
    GROUP BY knowbit_id
) AS SubTable
";

    $rs_max_headers = mysql_query($sql_max_headers, $db_conn);
    $row_rs_max_headers = mysql_fetch_assoc($rs_max_headers);

    for ($i = 1; $i <= $row_rs_max_headers['total']; $i++) {
        $comment_headers .= '"csv_comment_'.$i.'_content",';
        $comment_headers .= '"csv_comment_'.$i.'_author",';
        $comment_headers .= '"csv_comment_'.$i.'_date",';
    }
    return $comment_headers;
}


######################################################################
# GET_COMMENTS
######################################################################
function get_comments($db_conn, $knowbit_id) {
$sql_comments_per_knowbit = "
SELECT comment,
             username,
             submitted_date
FROM comments
WHERE username != 'will' AND
        knowbit_id = $knowbit_id";

    $rs_comments = mysql_query($sql_comments_per_knowbit, $db_conn);
    $row_rs_comments = mysql_fetch_assoc($rs_comments);
    $totalRows_rs_comments = mysql_num_rows($rs_comments);

    $comments = '';
    do {
        $content = add_extra_quotes(transcribe_cp1252_to_latin1($row_rs_comments['comment']));
        $author = $row_rs_comments['username'];
        $date = $row_rs_comments['submitted_date'];

        $comments .= '"'.$content.'", ';
        $comments .= '"'.$author.'", ';
        $comments .= '"'.$date.'", ';

    } while ($row_rs_comments = mysql_fetch_assoc($rs_comments));

    return $comments;
}


######################################################################
# GET_CATEGORIES
######################################################################
function get_categories($db_conn, $knowbit_id) {
$sql_categories_for_knowbit = "
SELECT cat.name AS name
FROM categories cat, assignments ass
WHERE cat.id = ass.category_id AND
          ass.knowbit_id = $knowbit_id";

    $rs_categories = mysql_query($sql_categories_for_knowbit, $db_conn);
    $row_rs_categories = mysql_fetch_assoc($rs_categories);
    $totalRows_rs_categories = mysql_num_rows($rs_categories);

    $categories = '"';

    do {
        $category = $row_rs_categories['name'];
        $categories .= $category.", ";
    } while ($row_rs_categories = mysql_fetch_assoc($rs_categories));
    $categories .= '", ';

    $fixed_categories = fix_categories($categories);

    return $fixed_categories;
}

######################################################################
# FIX_CATEGORIES : flip comma'd cats to slashed
######################################################################
function fix_categories($categories) {
    $patterns = array();
    $replacements = array();

    $patterns[0] = '/Div 06--Wood, Plastics, & Composites/';
    $replacements[0] = 'Div 06--Wood / Plastics / Composites';

    $patterns[3] = '/Flooring, Wood/';
    $replacements[3] = 'Flooring - Wood';

    $patterns[5] = '/Flooring, stone and tile/';
    $replacements[5] = 'Flooring - Stone & Tile';

    $patterns[1] = '/Insulation, vapor retarders, air barriers/';
    $replacements[1] = 'Insulation / Vapor Retarders / Air Barriers';

    $patterns[4] = '/Roofing, Green (planted)/';
    $replacements[4] = 'Roofing - Green (planted)';

    $patterns[2] = '/Toilet, bath & laundry accessories/';
    $replacements[2] = 'Toilet / Bath / Laundry Accessories';

    ksort($patterns);
    ksort($replacements);

    $fixed_categories = preg_replace($patterns, $replacements, $categories);

    return $fixed_categories;

}


######################################################################
# ADD_EXTRA_QUOTES
######################################################################
function add_extra_quotes($text) {
    $pattern = '/"/';
    $replacement = '""';
    $text = preg_replace($pattern, $replacement, $text);
    return $text;
}


######################################################################
# MERGE_INTO_POST
######################################################################
function merge_into_post($abstract, $description, $location, $phone, $url) {
    $post = "<strong>abstract:</strong> $abstract <br /><br />";
    $post .= "<strong>description:</strong> $description <br /><br />";
    if ($location != '') {
        $post .= "<strong>more:</strong> $location <br /><br />";
    }
    if ($phone != '') {
        $post .= "<strong>phone:</strong> $phone <br /><br />";
    }
    if ($url != '') {
        $post .= "<strong>url:</strong> $url <br /><br />";
    }

    // fix windows newlines
    $new_post = preg_replace('/(\r\n|\r|\n)/s',"\n",$post);

    $new_post = transcribe_cp1252_to_latin1($new_post);

    return $new_post;
}

######################################################################
# TRANSCRIBE_CP1252_TO_LATIN1: fix 'smart quotes' and the like (from http://php.net/manual/en/function.strtr.php)
######################################################################
function transcribe_cp1252_to_latin1($cp1252) {
  return strtr(
    $cp1252,
    array(
      "\x80" => "e",  "\x81" => " ",    "\x82" => "'", "\x83" => 'f',
      "\x84" => '"',  "\x85" => "...",  "\x86" => "+", "\x87" => "#",
      "\x88" => "^",  "\x89" => "0/00", "\x8A" => "S", "\x8B" => "<",
      "\x8C" => "OE", "\x8D" => " ",    "\x8E" => "Z", "\x8F" => " ",
      "\x90" => " ",  "\x91" => "`",    "\x92" => "'", "\x93" => '"',
      "\x94" => '"',  "\x95" => "*",    "\x96" => "-", "\x97" => "--",
      "\x98" => "~",  "\x99" => "(TM)", "\x9A" => "s", "\x9B" => ">",
      "\x9C" => "oe", "\x9D" => " ",    "\x9E" => "z", "\x9F" => "Y"));
}

?>
