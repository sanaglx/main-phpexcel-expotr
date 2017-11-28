<?php
/*
  Plugin Name: Main PHPExcel Export
  Description: Main PHPExcel Export Plugin for WordPress
  Version: 1.11
  Author: Sanagl
  Author URI: 
 */

define("SPEE_PLUGIN_URL", WP_PLUGIN_URL.'/'.basename(dirname(__FILE__)));
define("SPEE_PLUGIN_DIR", WP_PLUGIN_DIR.'/'.basename(dirname(__FILE__)));

add_action ( 'admin_menu', 'spee_admin_menu' );

function spee_admin_menu() {
	add_menu_page ( 'PHPExcel Export', 'Export', 'manage_options', 'spee-dashboard', 'spee_dashboard' );
}

function spee_dashboard() {
	global $wpdb;
	if ( isset( $_GET['export'] )) {
		if ( file_exists(SPEE_PLUGIN_DIR . '/lib/PHPExcel.php') ) {
			
			//Include PHPExcel
			require_once (SPEE_PLUGIN_DIR . "/lib/PHPExcel.php");
			
			// Create new PHPExcel object
			$objPHPExcel = new PHPExcel();
			
			// Set document properties
			
			// Add some data
			$objPHPExcel->setActiveSheetIndex(0);
			$objPHPExcel->getActiveSheet()->setCellValue('A1', 'Номер п/п');
			$objPHPExcel->getActiveSheet()->setCellValue('B1', 'Наименование товара');
			$objPHPExcel->getActiveSheet()->setCellValue('C1', 'Цена');
			$objPHPExcel->getActiveSheet()->setCellValue('D1', 'Цена со скидкой');
			$objPHPExcel->getActiveSheet()->setCellValue('E1', 'Каталог');
			//$objPHPExcel->getActiveSheet()->setCellValue('F1', 'Content');
			
			
			$objPHPExcel->getActiveSheet()->getStyle('A1:G1')->getFont()->setBold(true);
			$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('A:G')->setAutoSize(true);
			

/*****************************************AND ( wp_term_relationships.term_taxonomy_id IN (".$cat."))****/
$query="SELECT   wp_posts.* FROM wp_posts
INNER JOIN wp_term_relationships ON (wp_posts.ID = wp_term_relationships.object_id)
INNER JOIN wp_postmeta ON ( wp_posts.ID = wp_postmeta.post_id ) WHERE 1=1 
 AND (( wp_postmeta.meta_key = '_visibility' AND CAST(wp_postmeta.meta_value AS CHAR)
 IN ('catalog','visible') ))
 AND wp_posts.post_type = 'product'
 AND ((wp_posts.post_status = 'publish'))
 GROUP BY wp_posts.ID  ORDER BY wp_posts.post_title ASC ";
 

 
$posts   = $wpdb->get_results($query);


		if ( $posts ) {
				foreach ( $posts as $i=>$post ) {
					 $categories = get_the_terms($post->ID , 'product_cat' );
					 $cc="";
					 foreach ($categories as $category) {
						$cc = $category->name." , ".$cc;
					}	
					$objPHPExcel->getActiveSheet()->setCellValue('A'.($i+2+$a), $i+2);
					$objPHPExcel->getActiveSheet()->setCellValue('B'.($i+2+$a), $post->post_title);
					$objPHPExcel->getActiveSheet()->setCellValue('C'.($i+2+$a), get_post_meta( $post->ID, '_regular_price', true));
					$objPHPExcel->getActiveSheet()->setCellValue('D'.($i+2+$a), get_post_meta( $post->ID, '_sale_price', true));
					$objPHPExcel->getActiveSheet()->setCellValue('E'.($i+2+$a), $cc);
					
	
				}
			}
		

			// Rename worksheet
			//$objPHPExcel->getActiveSheet()->setTitle('Simple');
			
			// Set active sheet index to the first sheet, so Excel opens this as the first sheet
			$objPHPExcel->setActiveSheetIndex(0);
			
			// Redirect output to a client’s web browser
			ob_clean();
			ob_start();
			switch ( $_GET['format'] ) {
				case 'csv':
					// Redirect output to a client’s web browser (CSV)
					header("Content-type: text/csv");
					header("Cache-Control: no-store, no-cache");
					header('Content-Disposition: attachment; filename="export.csv"');
					$objWriter = new PHPExcel_Writer_CSV($objPHPExcel);
					$objWriter->setDelimiter(',');
					$objWriter->setEnclosure('"');
					$objWriter->setLineEnding("\r\n");
					//$objWriter->setUseBOM(true);
					$objWriter->setSheetIndex(0);
					$objWriter->save('php://output');
					break;
				case 'xls':
					// Redirect output to a client’s web browser (Excel5)
					header('Content-Type: application/vnd.ms-excel');
					header('Content-Disposition: attachment;filename="export.xls"');
					header('Cache-Control: max-age=0');
					$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
					$objWriter->save('php://output');
					break;
				case 'xlsx':
					// Redirect output to a client’s web browser (Excel2007)
					header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
					header('Content-Disposition: attachment;filename="export.xlsx"');
					header('Cache-Control: max-age=0');
					$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
					$objWriter->save('php://output');
					break;
			}
			exit;
		}
	} 
?>
<div class="wrap">
	<h2><?php _e( "Экспорт Прайс листа " ); ?></h2>
	<form method='get' action="admin.php?page=spee-dashboard">
		<input type="hidden" name='page' value="spee-dashboard"/>
		<input type="hidden" name='noheader' value="1"/>
		<input type="radio" name='format' id="formatCSV"  value="csv" checked="checked"/>  <label for"formatCSV">csv</label>
		<input type="radio" name='format' id="formatXLS"  value="xls"/>  <label for"formatXLS">xls</label>
		<input type="radio" name='format' id="formatXLSX" value="xlsx"/> <label for"formatXLSX">xslx</label>
		<input type="submit" name='export' id="csvExport" value="Export"/>
	</form>
	<div class="footer-credit alignright">
		<p>Source library based <a title="anang pratika" href="https://github.com/PHPOffice/PHPExcel<" target="_blank" >PHPExcel</a>.</p>
	</div>
</div>
<?php
}
