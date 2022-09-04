<?php
/* 
 /* Copyright (C) 2022
 * Author: Luigi Grillo - luigi.grillo@gmail.com (http://luigigrillo.com)
 * Creation date: 05/01/2022
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 * 
*/

/* TODO:
 *  - The parameter are taken from Order extrafields
 *  - The CMR Items are strictly defined according to the DOLIFarm.PriceGenerator managment. 
 *    It should be generalized.
 */

/**
 *	\file       dolicmr/index.php
 *	\ingroup    dolicmr
 *	\brief      Home page of dolicmr top menu
 */

$DEBUG = 0;

/*
set_error_handler(function(int $number, string $message) {
        GLOBAL $DEBUG;
   if ($DEBUG) {
           echo "Handler captured error $number: '$message'" . PHP_EOL  ;
   }
});
*/


// Load Dolibarr environment
$res = 0;
// Try main.inc.php into web root known defined into CONTEXT_DOCUMENT_ROOT (not always defined)
if (!$res && !empty($_SERVER["CONTEXT_DOCUMENT_ROOT"])) {
	$res = @include $_SERVER["CONTEXT_DOCUMENT_ROOT"]."/main.inc.php";
}
// Try main.inc.php into web root detected using web root calculated from SCRIPT_FILENAME
$tmp = empty($_SERVER['SCRIPT_FILENAME']) ? '' : $_SERVER['SCRIPT_FILENAME']; $tmp2 = realpath(__FILE__); $i = strlen($tmp) - 1; $j = strlen($tmp2) - 1;
while ($i > 0 && $j > 0 && isset($tmp[$i]) && isset($tmp2[$j]) && $tmp[$i] == $tmp2[$j]) {
	$i--; $j--;
}
if (!$res && $i > 0 && file_exists(substr($tmp, 0, ($i + 1))."/main.inc.php")) {
	$res = @include substr($tmp, 0, ($i + 1))."/main.inc.php";
}
if (!$res && $i > 0 && file_exists(dirname(substr($tmp, 0, ($i + 1)))."/main.inc.php")) {
	$res = @include dirname(substr($tmp, 0, ($i + 1)))."/main.inc.php";
}
// Try main.inc.php using relative path
if (!$res && file_exists("../main.inc.php")) {
	$res = @include "../main.inc.php";
}
if (!$res && file_exists("../../main.inc.php")) {
	$res = @include "../../main.inc.php";
}
if (!$res && file_exists("../../../main.inc.php")) {
	$res = @include "../../../main.inc.php";
}
if (!$res) {
	die("Include of main fails");
}

require_once DOL_DOCUMENT_ROOT.'/core/class/html.formfile.class.php';
require_once DOL_DOCUMENT_ROOT.'/core/class/html.formfile.class.php';

require_once DOL_DOCUMENT_ROOT.'/core/modules/commande/modules_commande.php';
require_once DOL_DOCUMENT_ROOT.'/commande/class/commande.class.php';
require_once DOL_DOCUMENT_ROOT.'/core/lib/order.lib.php';

require_once DOL_DOCUMENT_ROOT.'/expedition/class/expedition.class.php';
require_once DOL_DOCUMENT_ROOT.'/core/modules/expedition/modules_expedition.php';
require_once DOL_DOCUMENT_ROOT.'/core/class/extrafields.class.php';
require_once DOL_DOCUMENT_ROOT.'/core/lib/date.lib.php';

// 26-01-2022
	require_once DOL_DOCUMENT_ROOT.'/core/lib/functions.lib.php';


// require_once DOL_DOCUMENT_ROOT.'/includes/phpoffice/phpspreadsheet/src/autoloader.php';
require_once DOL_DOCUMENT_ROOT.'/includes/phpoffice/phpspreadsheet/src/autoloader.php';
require_once DOL_DOCUMENT_ROOT.'/includes/Psr/autoloader.php';
require_once PHPEXCELNEW_PATH.'Spreadsheet.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;




// Load translation files required by the page
$langs->loadLangs(array("dolicmr@dolicmr"));


// Security check
// if (! $user->rights->cmr->myobject->read) {
// 	accessforbidden();
// }
$socid = GETPOST('socid', 'int');
if (isset($user->socid) && $user->socid > 0) {
	$action = '';
	$socid = $user->socid;
}

$max = 5;
$now = dol_now();

/*
 * Actions
 */

// None


/*
 * View
 */
 
$form = new Form($db);
$formfile = new FormFile($db);

$orderObj = new Commande($db);
$societe = new Societe($db);
$courier =  new Societe($db);
$shipment = new Expedition($db);

$id=GETPOST('id','int');
$ref=GETPOST('ref','alpha');
$action = GETPOST('action', 'aZ09');
$palletType = GETPOST('palletType', 'aZ09');



//  Load Order Information
	$result = $orderObj->fetch($id);
	$orderObj->fetch_lines();
	$nOfLines = $orderObj->getNbOfProductsLines ();

	if ($DEBUG) echo "<br>number of products lines: $nOfLines<br>";


//  Getting extrafieds 
    $extrafields = new ExtraFields($db);
    $extrafields->fetch_name_optionals_label($orderObj->table_element);


    $dataPartenza = $orderObj->array_options["options_shipmentdate"];
	$luogoScarico = $orderObj->array_options["options_deliveryaddress"];
	$totalePedaneINDU = $orderObj->array_options["options_numindupallet"];
	$totalePedaneEURO = $orderObj->array_options["options_numeuropallet"];
	$luogoPartenza = $orderObj->array_options["options_shipmentaddress"];

    $error = 0;
	$sql= "SELECT libelle  FROM ".MAIN_DB_PREFIX."c_shipment_mode WHERE  rowid=$orderObj->shipping_method_id";
    $res=$db->query($sql);
	if ($res) {
		$obj = $db->fetch_object($res);
		if ($obj) {
			$emailCourier = $obj->libelle;
			$courier->fetch('','','','','','','','','','',$emailCourier,'');       
		} else {
			$error++;
			setEventMessages($obj->error, $obj->errors, 'errors');
		}
	} else {
		$error++;
		setEventMessages($langs->trans("ShipmentModeSelectError"), '', 'errors');
	}

	
	if (empty($action)) {
		llxHeader("", $langs->trans("CMRArea"));
		$head = commande_prepare_head($orderObj, $user) ;
		dol_fiche_head($head, $active='3', $title='', $notab=0, $picto='infobox-commande');
		print load_fiche_titre($langs->trans("CMRArea"), '', 'cmr.png@cmr');
		print '<div class="fichecenter"><div class="fichethirdleft">';
		echo "<br><br>";
		print '<br><br><form action="'.$_SERVER['PHP_SELF'].'" method="POST">';
		print '<input type="hidden" name="token" value="'.newToken().'">';
		print '<input type="hidden" name="action" value="create">';
		print '<input type="hidden" name="id" value="'.$id.'">';
		print '<input class="butAction" type="submit" value="'.$langs->trans("GenerateCMR").'">';
		print "</form>";
	}

if ($action == "create") {
   $error = 0;
// Loading XLX CMR Template
	$templateFileName = "templates/CMR_template.xls";
	$cmrSpreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($templateFileName);
    $sheet = $cmrSpreadsheet->getActiveSheet();

// 26-01-2022 I do not rely on note_public anymore
// Getting note_public information

    if ($DEBUG) echo "public note:<br>".$orderObj->note_public;
// Setting CMR number
	$sheet->setCellValue('T10', str_replace("CO","",$orderObj->ref));

// Setting courier
	 // print_r($courier);
	$sheet->setCellValue('M20',$courier->nom);
	$sheet->setCellValue('M21',$courier->address.' '.$courier->zip.' '.$courier->town.' ('.$courier->state_code.')');
	$sheet->setCellValue('M22',$courier->note_public); 
	$sheet->setCellValue('M23',$langs->trans("TVA").": ".$courier->tva_intra);
	
// Setting DESTINATARIO
	$societe->fetch($orderObj->socid);
	$sheet->mergeCells('E19:L19');
	$sheet->setCellValue('E19', $societe->name);
	
	if ($societe->eori != '') {
		$sheet->setCellValue('H20', $langs->trans("TVA").": $societe->tva_intra EORI:  $societe->eori ");
	} else {
		$sheet->setCellValue('H20', $langs->trans("TVA").": $societe->tva_intra");
	}
	$sheet->setCellValue('H21', $langs->trans("Tel").": ".$societe->phone);
	$sheet->setCellValue('H22', $societe->address);
	$sheet->setCellValue('H23', $societe->zip." ".$societe->town." ".$societe->state);

// Setting Luogo di Consegna
	$sheet->mergeCells('F28:K28');
	$sheet->setCellValue('F28', $societe->name);
	$sheet->mergeCells('F29:K29');
	$sheet->setCellValue('F29', $langs->trans("Tel").": ".$societe->phone);
	$sheet->mergeCells('E30:L31');
	$sheet->getStyle('E30:L31')->getAlignment()->setWrapText(true);
	if ($luogoScarico == '') {
		$sheet->setCellValue('E30', $societe->address.' '.$societe->zip." ".$societe->town." ".$societe->state);
		
	} else {
		$sheet->setCellValue('E30', $luogoScarico); 
	}
	$sheet->getStyle("E28:L31")->getFont()->setSize(14);
	$sheet->getStyle('E28:L31')->getAlignment()->setHorizontal('center');
	
// Setting Luogo e data della presa in Carico
	$sheet->mergeCells('F35:K35');
	$sheet->setCellValue('F35',$langs->trans("Shipment").': '.dol_print_date(strip_tags($dataPartenza),"%A %d/%m/%Y").$langs->trans("At").': ');
	$sheet->getStyle('F35')->getAlignment()->setWrapText(true);
	$sheet->mergeCells('F36:K37');
	$sheet->setCellValue('F36',strip_tags($luogoPartenza));
	$sheet->getStyle('F36')->getAlignment()->setWrapText(true);
	$sheet->getStyle("F35:K37")->getFont()->setSize(14);

	$sheet->setCellValue('H39',$langs->trans("ExtraNote"));

// Setting Documenti allegati
	$sql = "SELECT fk_target from ".MAIN_DB_PREFIX."element_element where fk_source=$orderObj->id";	
	$result=$db->query($sql);
	if (!$result) {
		$error++;
		setEventMessages($langs->trans("ShipmentDoesNotExist"), '', 'errors');
	} else $row = $db->fetch_array($result);
	$expID = $row[0];
	$sql = "SELECT ref, date_valid from ".MAIN_DB_PREFIX."expedition where rowid=$expID";
	$result=$db->query($sql);
	if (!$result) {
		$error++;
		setEventMessages($langs->trans("ShipmentDoesNotExist"), '', 'errors');
	} else $row = $db->fetch_array($result);
	$sheet->setCellValue('H44',$row['red']);
	$dateShipment = dol_print_date(dol_stringtotime($row['date_valid']),"%d-%m-%Y");
	$sheet->setCellValue('J44',$dateShipment);

// Insert product information
    if(!$error) {
		$qtyTotal = 0; 
		$colliTotal = 0;
		$pesoLordoTotale = 0;
		for ($o=0;$o<$nOfLines;$o++) {     // TODO CRITICAL: THe CMR ITEM are created from the Products label description
			 $l = get_object_vars($orderObj->lines[$o]);

			 $info = explode("<br />",$l["desc"]);
			 $denominazione = explode("|",$info[0]);   
			 $qty = explode("=",$info[1]);
			 $colli = explode("X",$qty[0]);
			 $p = explode("(",$qty[1]);
			 $pesoLordo=explode("kg",$p[1]);
			 $sheet->insertNewRowBefore(50+$o);
			 $sheet->setCellValue('H'.(50+$o), $colli[0]);
			 $sheet->setCellValue('J'.(50+$o), $qty[0]);
			 $sheet->getStyle('H'.(50+$o))->getFont()->setSize(14);
			 $sheet->setCellValue('L'.(50+$o), str_replace("&nbsp;","",$denominazione[1]));
				$sheet->getStyle('L'.(50+$o))->getFont()->setSize(14);

			 $sheet->setCellValue('Q'.(50+$o), $l['qty'].'Kg');
				$sheet->getStyle('Q'.(50+$o))->getFont()->setSize(14);
				
			$sheet->setCellValue('T'.(50+$o), $pesoLordo[0].'Kg');
			$sheet->getStyle('T'.(50+$o))->getFont()->setSize(14);

			 $qtyTotal = $qtyTotal+$l['qty'];
			 $colliTotal = $colliTotal + (int)$colli[0];
			 $pesoLordoTotale = $pesoLordoTotale + $pesoLordo[0];
	}
		$sheet->setCellValue('Q'.(50+$o+1), $qtyTotal.'Kg');
		$sheet->setCellValue('H'.(50+$o+1), $colliTotal);
		$sheet->setCellValue('T'.(50+$o+1), $pesoLordoTotale.'Kg');
		$sheet->getStyle('Q'.(50+$o+1))->getFont()->setSize(14);
		$sheet->getStyle('H'.(50+$o+1))->getFont()->setSize(14);
		$sheet->getStyle('T'.(50+$o+1))->getFont()->setSize(14);
		
	// Setting Total pallets
		if ($totalePedaneINDU != '')
			$sheet->setCellValue('P'.(50+$o+3),$langs->trans('numberOfINDUPallet').': '.strip_tags($totalePedaneINDU));
		else 
			$sheet->setCellValue('P'.(50+$o+3),$langs->trans('numberOfEUROPallet').': '.strip_tags($totalePedaneEURO));
		
		$sheet->getStyle('P'.(50+$o+3))->getFont()->setSize(14);

	// Setting Totale Incoterm
		$sql = "SELECT code from ".MAIN_DB_PREFIX."c_incoterms where rowid=$societe->fk_incoterms";
		$result=$db->query($sql);
		if (!$result) {
			$error++;
			setEventMessages($langs->trans("c_incotermsDoesNotExist"), '', 'errors');
		} else $row = $db->fetch_array($result);
		$sheet->setCellValue('E'.(50+$o+8),$langs->trans('Incoterm').": ".$row[0]." - ".$langs->trans('ORIGINEITALIA')); // TODO Origine Italia !!
		$sheet->getStyle('E'.(50+$o+8))->getFont()->setSize(14);

	// Creating CMR xls file   // TODO It is very speific to the Mandala Case
		$ddate = date('Y-m-d');
		$date = new DateTime($ddate);
		$fileName = dol_sanitizeFileName(str_replace('#','',$orderObj->ref)).'/'."CMR".str_replace('CO','',$orderObj->ref)."_".$societe->name."_".dol_sanitizeFileName(str_replace('#','',$orderObj->ref_client))."_".$dateShipment.".xlsx";
		$cmrOutput = $conf->commande->multidir_output[$orderObj->entity].'/'.$fileName;
		$writer = new Xlsx($cmrSpreadsheet);
		$writer->save($cmrOutput);

	// Creating link to the CMR xls file
		// echo '<br> <a href="'.DOL_URL_ROOT.'/document.php?modulepart=commande&entity=1&file='.$fileName.'">'.$langs->trans("CMRGenerated").'</a>' ;
		// echo "<br><br>DONE!";	
   } // END if(!$error)
	header("Location: ".DOL_URL_ROOT.'/commande/card.php?id='.$id);
} // END if($action == 'create')

// print '</div><div class="fichetwothirdright"><div class="ficheaddleft">';

if (empty($action)) {
	print '</div></div></div>';
	// End of page
	llxFooter();
} else {
	// dol_htmloutput_events(1);
}

$db->close();
