/***
	
	Premier essai

***/


#target photoshop
#include "TableauExifs.jsx"
#include "FonctionsCommunes.jsx"

function ProgrammePrincipal () {
	$.level = 1;
	this.elementsRequis="Il faut démarrer photoshop avec ExtendedToolkit";
	// load the library
	if ( ExternalObject.AdobeXMPScript == undefined ) {
		ExternalObject.AdobeXMPScript = new ExternalObject( "lib:AdobeXMPScript");
		$.writeln("lib:AdobeXMPScript chargée");
	}
	app.preferences.rulerUnits=Units.PIXELS


}

ProgrammePrincipal.prototype.redimensionne = function(photo) {	

	try {
		open (photo);
		var fichierOuvert = activeDocument;
		var hauteur=fichierOuvert.height;
		var largeur=fichierOuvert.width;
		var horizontal=true;
		if (hauteur>largeur) {
			var horizontal=false;
// 1920x1080 Vertical
			var hauteur10x15=1920;
			var largeur10x15=hauteur10x15*largeur/hauteur;
			if (largeur10x15<1080) {
				var hauteur10x15=hauteur10x15*1080/largeur10x15;
				var largeur10x15=1080;
			}
		} else {

// 1920x1080 Horizontal
			var largeur10x15=1920;
			var hauteur10x15=largeur10x15*hauteur/largeur;
			if (hauteur10x15<1080) {
				var largeur10x15=largeur10x15*1080/hauteur10x15;
				var hauteur10x15=1080;
			}
		}

// Préparation de la photo 10x15
//~ 		$.writeln(largeur10x15 + " x " + hauteur10x15);
		var sauve=false;
		fichierOuvert.resizeImage(largeur10x15,hauteur10x15,72);
		if (horizontal) {
			if (largeur!=1920||hauteur!=1080) {
				fichierOuvert.resizeCanvas(1920,1080,AnchorPosition.MIDDLECENTER);
				sauve=true;
			}
		} else {
			if (largeur!=1080||hauteur!=1920) {
				fichierOuvert.resizeCanvas(1080,1920,AnchorPosition.MIDDLECENTER);
				sauve=true;
			}
		}

		if (sauve) {
			fichierOuvert.save();
		}
		fichierOuvert.close(SaveOptions.DONOTSAVECHANGES);
	}
	catch(x){};
	return true;
}


ProgrammePrincipal.prototype.toutEtOK = function() {	
	// Photoshop must be running
	if (BridgeTalk.isRunning("photoshop")) {
		$.writeln("OK, photoshop tourne !");
		return true;		
	}
	
	// Fail if these preconditions are not met 
	$.writeln("ERREUR: Ne peut pas démarrer le programme principal...");
	$.writeln(this.elementsRequis);
	return false;	
}


ProgrammePrincipal.prototype.run = function() {
	
	if (!this.toutEtOK()) {
		$.writeln ("Ca ne marche pas");
		return false;
	}

	
	$.writeln("OK, ça tourne !");
	$.getenv("temp");
	
// Recherche du dernier chemin parcouru

	var dernierRepertoireLu="~";
	var fichierDeConfiguration=File("~/dernierRepertoireLu.txt");
	if (fichierDeConfiguration.exists) {
		fichierDeConfiguration.open("r");
		while (!fichierDeConfiguration.eof) {
			var dernierRepertoireLu=fichierDeConfiguration.readln();
		}
		fichierDeConfiguration.close();
	}

	var repertoire=Folder.selectDialog ("Choix du répertoire :",dernierRepertoireLu);
	
	$.writeln("Chemin : " + repertoire.path + "/" + repertoire.name);
	
// Stockage du dernier répertoire lu

	fichierDeConfiguration.open("w");
	fichierDeConfiguration.writeln(repertoire.path + "/" + repertoire.name);
	fichierDeConfiguration.close();
	$.writeln("Fichier : " + fichierDeConfiguration.path + "/" + fichierDeConfiguration.name);


	this.traitementDesPhotos(repertoire);
	
	
	return true;

}

ProgrammePrincipal.prototype.traitementDesPhotos = function(repertoire) {


	var liste=repertoire.getFiles("*.jpg");
	$.writeln("#### traitementDesPhotos #####" + repertoire.fullName + "/" + repertoire.name);
	$.writeln("Nombre de photos : " + liste.length);

	ProgrammePrincipal.tableauFormatHTML=new File(repertoire.fullName+"/"+repertoire.name+".html");
	ProgrammePrincipal.tableauFormatHTML.open("w");
	ProgrammePrincipal.tableauFormatHTML.writeln("<!DOCTYPE html>");
	ProgrammePrincipal.tableauFormatHTML.writeln("<HTML>");
	ProgrammePrincipal.tableauFormatHTML.writeln("<HEAD>");
	ProgrammePrincipal.tableauFormatHTML.writeln("</HEAD>");
	ProgrammePrincipal.tableauFormatHTML.writeln("<BODY>");

	for (var p=0;p<liste.length;p++) {
		var photo=liste[p];
		if (photo instanceof File) {
			$.writeln();
//~ 			$.writeln("Poids de la photo " + p + " : " + photo.length);
//~ 			try {

				var tableauDesValeursXMP=this.donneesXMP(photo);
//~ 				$.writeln(tableauDesValeursXMP);
				var DateTimeOriginal=tableauDesValeursXMP["exif"]["DateTimeOriginal"];
				var PixelXDimension=tableauDesValeursXMP["exif"]["PixelXDimension"];
				var PixelYDimension=tableauDesValeursXMP["exif"]["PixelYDimension"];
//~ 				$.writeln(DateTimeOriginal + " " + PixelXDimension + " " + PixelYDimension);
				var DateTimeOriginal=String(DateTimeOriginal).split("T");
				var annee=DateTimeOriginal[0].split("-");
				var mois=annee[1];
				var jour=annee[2];
				var annee=annee[0];
				var heure=DateTimeOriginal[1]!=undefined?DateTimeOriginal[1].split(":"):new Array("00","00","00.00");
				var minutes=heure[1];
				var secondes=heure[2]!=undefined?heure[2].split("."):new Array("00","00");
				var millisecondes=secondes[1]!=undefined?secondes[1]:"00";
//~ 				$.writeln("millisecondes "+ millisecondes);
				var secondes=secondes[0];
				var heure=heure[0];
				var DateTimeOriginal=String(annee) + mois + jour + heure + minutes + secondes;
	
// Renomme la photo
				var compteur=parseInt(millisecondes);
				var leFichierExiste=true;
				while (leFichierExiste) {
					var numero="0" + compteur++;
					var numero=numero.substring(numero.length - 2);
					var nouveauNomDeLaPhoto=DateTimeOriginal + numero + ".jpg";
					var leFichierExiste=File(repertoire.fullName + "/" + nouveauNomDeLaPhoto).exists;
				}

				$.writeln("Nom de la photo " + p + " : " + photo.name + ", nouveau nom : " + nouveauNomDeLaPhoto);
				if (photo.name!=nouveauNomDeLaPhoto) {
					photo.rename(nouveauNomDeLaPhoto);
					$.writeln("Renommée "  + nouveauNomDeLaPhoto);
				}
				
				

				ProgrammePrincipal.tableauFormatHTML.writeln("<H3>");
				ProgrammePrincipal.tableauFormatHTML.writeln(photo.name);
				ProgrammePrincipal.tableauFormatHTML.writeln("</H3>");
				ProgrammePrincipal.tableauFormatHTML.writeln("<IMG src=\""+photo.name+"\" width=\"500px\" />");
				ProgrammePrincipal.tableauFormatHTML.writeln("<table border=1>");
				ProgrammePrincipal.tableauFormatHTML.writeln("<TR>");
				ProgrammePrincipal.tableauFormatHTML.writeln("<TD>GROUPE</TD><TD>TAG</TD><TD>Valeur</TD>");
				ProgrammePrincipal.tableauFormatHTML.writeln("</TR>");
				for (var x in tableauDesValeursXMP) {
					for (var t in tableauDesValeursXMP[x]) {
						ProgrammePrincipal.tableauFormatHTML.writeln("<TR>");
						ProgrammePrincipal.tableauFormatHTML.writeln("<TD>"+x+"</TD><TD>"+t+"</TD><TD>"+tableauDesValeursXMP[x][t]+"</TD>");
						ProgrammePrincipal.tableauFormatHTML.writeln("</TR>");
					}
				}
//~ 				ProgrammePrincipal.tableauFormatHTML.writeln("</table>");

					var tableauDesDonneesExif=this.informationsSurLimage(photo);
//~ 				ProgrammePrincipal.tableauFormatHTML.writeln("<table border=1>");
//~ 				ProgrammePrincipal.tableauFormatHTML.writeln("<TR>");
//~ 				ProgrammePrincipal.tableauFormatHTML.writeln("<TD>TAG</TD><TD>Valeur</TD>");
//~ 				ProgrammePrincipal.tableauFormatHTML.writeln("</TR>");
				for (var x in tableauDesDonneesExif) {
					ProgrammePrincipal.tableauFormatHTML.writeln("<TR>");
//~ 			$.writeln(x + " = " + tableauDesDonneesExif[x]);
					ProgrammePrincipal.tableauFormatHTML.writeln("<TD>exif</TD><TD>"+x+"</TD><TD>"+tableauDesDonneesExif[x]+"</TD>");
					ProgrammePrincipal.tableauFormatHTML.writeln("</TR>");
				}
//~ 				ProgrammePrincipal.tableauFormatHTML.writeln("</table>");
/***
				if (PixelXDimension!=1920||PixelYDimension!=1080) {
					this.redimensionne(photo);
				}
***/
/***
	}
			catch(e) {
				$.writeln("Erreur " + e);
			}
***/
		}
	}
		
	ProgrammePrincipal.tableauFormatHTML.writeln("</BODY>");
	ProgrammePrincipal.tableauFormatHTML.writeln("</HTML>");
	ProgrammePrincipal.tableauFormatHTML.close();
	return true;
	
}

ProgrammePrincipal.prototype.informationsSurLimage = function(photo) {

	var tableauDesDonneesExif=new Array();
	var tableauDesTagsExif=CreationDuTableauDesTagsExif();
	photoshop.open(photo);

	for (var i in activeDocument.info) {
//~ 		$.writeln(i+ " --- " + Object.prototype.toString.call(activeDocument.info[i]));
		if (Object.prototype.toString.call(activeDocument.info[i])=="[object Array]") {
			if (i!="exif") {
				var tableau=activeDocument.info[i];
				for (var n in tableau) {
//~ 					$.writeln(i + "[" + n + "] = " + tableau[n]);
						tableauDesDonneesExif[i + "/" +n]=tableau[n];
				}
			}
		} else if (Object.prototype.toString.call(activeDocument.info[i])=="[object String]") {
			tableauDesDonneesExif[i]=activeDocument.info[i];
//~ 			$.writeln(i + " = " + activeDocument.info[i]);
		}
	}


	var donneesEXIF=activeDocument.info.exif;
	for (var x in donneesEXIF) {
		var tableauDesValeurs=donneesEXIF[x];
//~ 		$.writeln ((tableauDesTagsExif[tableauDesValeurs[2]]!=undefined?tableauDesValeurs[2] + ":" + tableauDesTagsExif[tableauDesValeurs[2]].label:tableauDesValeurs[0]) + " = " + tableauDesValeurs[1])
			tableauDesDonneesExif[(tableauDesTagsExif[tableauDesValeurs[2]]!=undefined?tableauDesTagsExif[tableauDesValeurs[2]].label:tableauDesValeurs[0])]=tableauDesValeurs[1];
	}
	activeDocument.close(SaveOptions.DONOTSAVECHANGES);
	return tableauDesDonneesExif;
}

ProgrammePrincipal.prototype.donneesXMP = function(photo) {

	var fichierXMP=new XMPFile(photo.fsName,XMPConst.UNKNOWN,XMPConst.OPEN_FOR_UPDATE);
	var tableauXMP=fichierXMP.getXMP();

// Mise à jour d'une donnée EXIF

//~ 	tableauXMP.deleteProperty("http://cipa.jp/exif/1.0/","BodySerialNumber");
//~ 	tableauXMP.appendArrayItem("http://cipa.jp/exif/1.0/", "BodySerialNumber", "***************" + Math.round(Math.random()*Math.pow (10,10)) + "***************", 0, XMPConst.ARRAY_IS_ORDERED );



	tableauXMP.deleteProperty(XMPConst.NS_DC,"subject");
	tableauXMP.deleteProperty(XMPConst.NS_DC,"identifier");
	tableauXMP.deleteProperty(XMPConst.NS_EXIF,"identifier");
	var identifiantExistant=tableauXMP.getProperty(XMPConst.NS_XMP,"identifier");
	if (identifiantExistant==undefined) {
		var maintenant=new Maintenant();
		var identifiant=maintenant.ilEstExactement;
		tableauXMP.setProperty(XMPConst.NS_XMP,"identifier",identifiant,0,XMPConst.STRING);
//~ 		$.writeln("Nouvel identifiant = " + identifiant);
	} else {
//~ 		$.writeln("Identifiant existant = " + identifiantExistant);
	}
/**
	tableauXMP.appendArrayItem(XMPConst.NS_DC, "subject", "Mot clé 10", 0, XMPConst.ARRAY_IS_ORDERED );
	tableauXMP.appendArrayItem(XMPConst.NS_DC, "subject", "Mot clé 20", 0, XMPConst.ARRAY_IS_ORDERED );
	tableauXMP.appendArrayItem(XMPConst.NS_DC, "subject", "Mot clé 30", 0, XMPConst.ARRAY_IS_ORDERED );
	tableauXMP.appendArrayItem(XMPConst.NS_DC, "subject", "Mot clé 40", 0, XMPConst.ARRAY_IS_ORDERED );
**/

//~ 	tableauXMP.deleteProperty("http://ns.microsoft.com/photo/1.0/","LastKeywordXMP");
//~ 	tableauXMP.deleteProperty("http://ns.microsoft.com/photo/1.0/","LastKeywordIPTC");

	tableauXMP.deleteProperty(XMPConst.NS_DC,"usercomment");
	tableauXMP.deleteProperty(XMPConst.NS_DC,"XPComment");

	tableauXMP.deleteProperty(XMPConst.NS_DC,"title");
//~ 	tableauXMP.setProperty( XMPConst.NS_DC, "title", "blablablabla", 0, XMPConst.STRING );

	tableauXMP.deleteProperty(XMPConst.NS_DC,"object");
	tableauXMP.deleteProperty(XMPConst.NS_DC,"XPSubject");
	tableauXMP.deleteProperty(XMPConst.NS_EXIF,"XPSubject");

	try {
		tableauXMP.deleteProperty("http://ns.microsoft.com/photo/1.0/","XPSubject");
	}
	catch(e) {
//~ 		$.writeln(e);
	}
	try {
		tableauXMP.deleteProperty("http://ns.microsoft.com/photo/1.0/","Object");
	}
	catch(e) {
//~ 		$.writeln(e);
	}
	try {
		tableauXMP.deleteProperty("http://ns.microsoft.com/photo/1.0/","LastObjectXMP");
	}
	catch(e) {
//~ 		$.writeln(e);
	}
//~ 	tableauXMP.setProperty("http://ns.microsoft.com/photo/1.0/", "LastObjectXMP", "gdgdfdfgdghshgfsdgh", 0, XMPConst.STRING );

//~ 	tableauXMP.deleteProperty(XMPConst.NS_DC,"description");
//~ 	tableauXMP.setProperty( XMPConst.NS_DC, "description", "blablablabla", 0, XMPConst.STRING );

	tableauXMP.deleteProperty(XMPConst.NS_DC,"imageuniqueid");
	tableauXMP.deleteProperty(XMPConst.NS_XMP,"imageuniqueid");
//~ 	tableauXMP.setProperty( XMPConst.NS_XMP, "imageuniqueid", "***************" + Math.round(Math.random()*Math.pow (10,10)) + "***************", 0, XMPConst.STRING );

	tableauXMP.deleteProperty(XMPConst.NS_XMP,"usercomment");
	tableauXMP.deleteProperty(XMPConst.NS_XMP,"XPComment");
//~ 	tableauXMP.setProperty(XMPConst.NS_XMP,"XPComment", "#################sdfsdfsdfsdfd################", 0, XMPConst.STRING );

	tableauXMP.deleteProperty(XMPConst.NS_EXIF,"XPComment");
//~ 	tableauXMP.setProperty(XMPConst.NS_EXIF,"XPComment", "#################sdfsdfsdfsdfd################", 0, XMPConst.STRING );

	
	// Write updated metadata into the file
	if ( fichierXMP.canPutXMP( tableauXMP ) ) {
		fichierXMP.putXMP( tableauXMP );
//~ 		$.writeln("Ajoute propriété");
	}
	fichierXMP.closeFile( XMPConst.CLOSE_UPDATE_SAFELY );

	var fichierXMP=new XMPFile(photo.fsName,XMPConst.UNKNOWN,XMPConst.OPEN_FOR_READ);
	var tableauXMP=fichierXMP.getXMP();
	var iterator=tableauXMP.iterator();
	var continu=true;

	var tableauDesValeursXMP=new Array();

	while (continu) {
		var suivant=iterator.next();
		if (suivant==void null) {
			break;
		} else {
			if (suivant.value!="") {
				var nomDuTag=suivant.path.split(":");
				if (tableauDesValeursXMP[nomDuTag[0]]==undefined) {
					tableauDesValeursXMP[nomDuTag[0]]=new Array();
				}
//~ 				$.writeln("[" + nomDuTag[0] + "]:[" + nomDuTag[1] + "]=" + suivant.value);
				tableauDesValeursXMP[nomDuTag[0]][nomDuTag[1]]=suivant.value;
/***
				switch (suivant.namespace) {
					case XMPConst.NS_DC :
						$.writeln("DC : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_IPTC_CORE :
						$.writeln("IPTC_CORE : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_RDF :
						$.writeln("RDF : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_XML :
						$.writeln("XML : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_XMP :
						$.writeln("XMP : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_XMP_RIGHTS :
						$.writeln("XMP_RIGHTS : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_XMP_MM :
						$.writeln("XMP_MM : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_XMP_BJ :
						$.writeln("XMP_BJ : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_XMP_NOTE :
						$.writeln("XMP_NOTE : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_PDF :
						$.writeln("PDF : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_PDFX :
						$.writeln("PDFX : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_PHOTOSHOP :
						$.writeln("PHOTOSHOP : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_PS_ALBUM :
						$.writeln("PS_ALBUM : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_EXIF :
						$.writeln("EXIF : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_EXIF_AUX :
						$.writeln("EXIF_AUX : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_TIFF :
						$.writeln("TIFF : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_PNG :
						$.writeln("PNG : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_JPEG :
						$.writeln("JPEG : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_SWF :
						$.writeln("SWF : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_JPK :
						$.writeln("JPK : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_CAMERA_RAW :
						$.writeln("CAMERA_RAW : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_DM :
						$.writeln("DM : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_ADOBE_STOCK_PHOTO :
						$.writeln("ADOBE_STOCK_PHOTO : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case XMPConst.NS_ASF :
						$.writeln("ASF : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case "http://cipa.jp/exif/1.0/" :
						$.writeln("ExifEX : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case "http://ns.microsoft.com/photo/1.0/" :
						$.writeln("MicrosoftPhoto : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					case "http://ns.adobe.com/photoshop/1.0/panorama-profile" :
						$.writeln("Panorama : "  + suivant.path + "=" + suivant.value + " " + suivant.options);
					break;
					default :
						$.writeln(suivant.namespace + " : " + suivant.path + "=" + suivant.value);
					break;
				}
***/
			}
		}
	}

	fichierXMP.closeFile();

//~ 	$.writeln(typeof(tableauDesValeursXMP));
	return tableauDesValeursXMP;
}

new ProgrammePrincipal().run();
