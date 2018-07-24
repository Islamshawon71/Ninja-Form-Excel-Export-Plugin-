<?php
/*
  Plugin Name: Simple PHPExcel Export
  Description: Simple PHPExcel Export Plugin for WordPress
  Version: 1.0.0
  Author: Mithun
  Author URI: http://twitter.com/mithunp
 */

define("SPEE_PLUGIN_URL", WP_PLUGIN_URL.'/'.basename(dirname(__FILE__)));
define("SPEE_PLUGIN_DIR", WP_PLUGIN_DIR.'/'.basename(dirname(__FILE__)));

add_action ( 'admin_menu', 'spee_admin_menu' );

function spee_admin_menu() {
	add_menu_page ( 'PHPExcel Export', 'Export', 'manage_options', 'spee-dashboard', 'spee_dashboard' );
}
add_action('wp_ajax_spee_dashboard', 'spee_dashboard');
add_action('wp_ajax_nopriv_spee_dashboard', 'spee_dashboard');


function spee_dashboard() {
	$date1 = $_REQUEST['date1'];
	$date2 = $_REQUEST['date2'];
	require_once 'lib/PHPExcel.php';
	$objPHPExcel = new PHPExcel();


	$fields_meta = Ninja_Forms()->form(22)->get_fields();

	$objPHPExcel->getActiveSheet()->setCellValue('A1', 'ID');
	$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Submission date');
	$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setBold(true);
	$col = 'C';
	$row = 1;
	foreach($fields_meta as $field){
		$field_settings = $field->get_settings();
		$field_names[$field_settings['key']] = sanitize_title( $field_settings['label'].'-'.$field_settings['key'] );
		$field_types[$field_settings['key']] = $field_settings['type'];

		if($field_settings['type'] != 'submit' && $field_settings['type'] != 'html'){

			$val = $selected_field_names[$field_settings['key']]=(isset($field_settings['admin_label']) && $field_settings['admin_label'] ? $field_settings['admin_label'] : $field_settings['label']);
			$objPHPExcel->getActiveSheet()->setCellValue($col.$row, $val);
			$objPHPExcel->getActiveSheet()->getStyle($col.$row)->getFont()->setBold(true);
			$col++;
		}

	}
	$row = 2;
	$submitions = get_submissions('22',$date1,$date2);
	$m = 1;
	$n = 0;
	foreach($submitions as $item){
 
		$objPHPExcel->getActiveSheet()->setCellValue('A'.$row,$m);
		$objPHPExcel->getActiveSheet()->setCellValue('B'.$row,$item['date_submitted']);
		$col = 'C';
		foreach ($fields_meta as $field) {
			$field_settings = $field->get_settings();
			$field_names[$field_settings['key']] = sanitize_title( $field_settings['label'].'-'.$field_settings['key'] );
			$field_types[$field_settings['key']] = $field_settings['type'];

			if($field_settings['type'] != 'submit' && $field_settings['type'] != 'html'){

				//$val = $selected_field_names[$field_settings['key']]=(isset($field_settings['admin_label']) && $field_settings['admin_label'] ? $field_settings['admin_label'] : $field_settings['label']);
				if( in_array($field_types[$field_settings['key']], array( 'listcheckbox' ) ) ){
					$field_output = '';
					$field_value = unserialize( $item[$field_settings['key']] );
					if( is_array($field_value) ){
						$field_output = '';
						foreach ($field_value as $key => $value) {
							if( $field_output == '' )
								$field_output = $value;
							else
								$field_output .= ', ' . $value;
						}
					}
					$item[$field_settings['key']] = $field_output;
				}
				if( in_array($field_types[$field_settings['key']], array( 'checkbox' ) ) ){
					if($item[$field_settings['key']] == 0){
						$item[$field_settings['key']] = 'No';
					}else{
						$item[$field_settings['key']] = 'Yes';
					}
				}
				if( in_array($field_types[$field_settings['key']], array( 'file_upload' ) ) ){

					$item[$field_settings['key']] =$item[$field_settings['key']][1];

				}
				//$objPHPExcel->getActiveSheet()->setCellValue($col.$row,$field_settings['key']);
				$objPHPExcel->getActiveSheet()->setCellValue($col.$row,$item[$field_settings['key']]);
				$col++;
			}
			//$objPHPExcel->getActiveSheet()->setCellValue($col.$row,$item[$field_settings['key']]);


		}
		$row++;
		$m++;

	}



	$objPHPExcel->getActiveSheet()->setTitle('Chesse1');
	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	header('Content-Disposition: attachment;filename="helloworld.xlsx"');
	header('Cache-Control: max-age=0');
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	ob_start();
	$objWriter->save("php://output");
	$xlsData = ob_get_contents();
	ob_end_clean();
 	$response =  array(
		'op' => 'ok',
		'file' => "data:application/vnd.ms-excel;base64,".base64_encode($xlsData)
	);

	die(json_encode($response));
}

function get_submissions($form_id,$begin_date,$end_date){
	$query_args = array(
		'post_type'         => 'nf_sub',
		'posts_per_page'    => 100,
		'date_query'        => array(
			'inclusive'     => true,
		),
		'meta_query'        => array(
			array(
				'key' => '_form_id',
				'value' => $form_id,
			)
		),
	);

	if( isset( $begin_date ) AND $begin_date != '') {
		$query_args['date_query']['after'] = $begin_date .= ' 00:00:00';
	}

	if( isset( $end_date ) AND $end_date != '' ) {
		$query_args['date_query']['before'] = $end_date .= ' 23:59:59';
	}

	$subs = new WP_Query( $query_args );

	$sub_objects = array();
	$sub_index = 0;

	if ( is_array( $subs->posts ) && ! empty( $subs->posts ) ) {
		foreach ( $subs->posts as $sub ) {
			$sub_objects[$sub_index] = Ninja_Forms()->form( $form_id )->get_sub( $sub->ID )->get_field_values();
			$sub_objects[$sub_index]['date_submitted'] = get_the_date('', $sub->ID );
			$sub_index++;
		}
	}

	return $sub_objects;
}

