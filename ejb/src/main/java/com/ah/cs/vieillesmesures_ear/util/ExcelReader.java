package com.ah.cs.vieillesmesures_ear.util;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.ah.cs.vieillesmesures_ear.exception.EnteteException;
import com.ah.cs.vieillesmesures_ear.exception.ExcelErrorException;
import com.ah.cs.vieillesmesures_ear.exception.NbrOngletException;

public class ExcelReader {

	private Workbook workbook;

	public static enum ENTETES {
		T("T"), NOM("NOM"), PRECISONSSYNONYME("Précisions/Synonyme"), NOMMULTIPLE("Nom  multiple"), RAP1(
				"RAP1"), SOUSMULT("SOUS-MULT,"), RAP2("RAP2"), VALEURDECIMALE("Valeur décimale"), EPOQUE("EPOQUE"), DPT(
						"DPT"), LIEUUSAGE("Lieu d'usage"), CANTONS1973("Cantons 1793"), CANTONS1801(
								"Cantons 1801/2013"), CR("C-R"), UTILISATION(
										"Utilisation"), SOURCESETCOMMENTAIRES("SOURCES et COMMENTAIRES");

		private final String text;

		private ENTETES(final String text) {
			this.text = text;
		}

		@Override
		public String toString() {
			return text;
		}

	};

	public ExcelReader(Path path)
			throws EncryptedDocumentException, InvalidFormatException, IOException, NbrOngletException {

		// Creating a Workbook from an Excel file (.xls or .xlsx)
		workbook = WorkbookFactory.create(path.toFile());

		if (workbook.getNumberOfSheets() > 1)
			throw new NbrOngletException("Fichier :" + path.getFileName() + " Nombre d'onglets supérieur à 1");

	}

	public ArrayList<ArrayList<String>> charger() throws EnteteException, ExcelErrorException, IOException {

		Sheet sheet = workbook.getSheetAt(0);

		// contrôle sur les entêtes de colonnes

		Iterator<Row> rowIterator = sheet.rowIterator();
		Row entete = rowIterator.next();

		Iterator<Cell> cellIterator = entete.cellIterator();

		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();

			switch (cell.getColumnIndex()) {
			case 0:
				if (!cell.getStringCellValue().equals(ENTETES.T.toString()))
					throw new EnteteException(ENTETES.T + " non trouvée");
				break;
			case 1:

				if (!cell.getStringCellValue().equals(ENTETES.NOM.toString()))
					throw new EnteteException(ENTETES.NOM + " non trouvée");
				break;
			case 2:
				if (!cell.getStringCellValue().equals(ENTETES.PRECISONSSYNONYME.toString()))
					throw new EnteteException(ENTETES.PRECISONSSYNONYME + " non trouvée");
				break;
			case 3:
				if (!cell.getStringCellValue().equals(ENTETES.NOMMULTIPLE.toString()))
					throw new EnteteException(ENTETES.NOMMULTIPLE + " non trouvée");
				break;
			case 4:
				if (!cell.getStringCellValue().equals(ENTETES.RAP1.toString()))
					throw new EnteteException(ENTETES.RAP1 + " non trouvée");
				break;
			case 5:
				if (!cell.getStringCellValue().equals(ENTETES.SOUSMULT.toString()))
					throw new EnteteException(ENTETES.SOUSMULT + " non trouvée");
				break;
			case 6:
				if (!cell.getStringCellValue().equals(ENTETES.RAP2.toString()))
					throw new EnteteException(ENTETES.RAP2 + " non trouvée");
				break;
			case 7:
				if (!cell.getStringCellValue().equals(ENTETES.VALEURDECIMALE.toString()))
					throw new EnteteException(ENTETES.VALEURDECIMALE + " non trouvée");
				break;
			case 8:
				if (!cell.getStringCellValue().equals(ENTETES.EPOQUE.toString()))
					throw new EnteteException(ENTETES.EPOQUE + " non trouvée");
				break;
			case 9:
				if (!cell.getStringCellValue().equals(ENTETES.DPT.toString()))
					throw new EnteteException(ENTETES.DPT + " non trouvée");
				break;
			case 10:
				if (!cell.getStringCellValue().equals(ENTETES.LIEUUSAGE.toString()))
					throw new EnteteException(ENTETES.LIEUUSAGE + " non trouvée");
				break;
			case 11:
				if (!cell.getStringCellValue().equals(ENTETES.CANTONS1973.toString()))
					throw new EnteteException(ENTETES.CANTONS1973 + " non trouvée");
				break;
			case 12:
				if (!cell.getStringCellValue().equals(ENTETES.CANTONS1801.toString()))
					throw new EnteteException(ENTETES.CANTONS1801 + " non trouvée");
				break;
			case 13:
				if (!cell.getStringCellValue().equals(ENTETES.CR.toString()))
					throw new EnteteException(ENTETES.CR + " non trouvée");
				break;
			case 14:
				if (!cell.getStringCellValue().equals(ENTETES.UTILISATION.toString()))
					throw new EnteteException(ENTETES.UTILISATION + " non trouvée");
				break;
			case 15:
				if (!cell.getStringCellValue().equals(ENTETES.SOURCESETCOMMENTAIRES.toString()))
					throw new EnteteException(ENTETES.SOURCESETCOMMENTAIRES + " non trouvée");
				break;
			}

		}

		ArrayList<ArrayList<String>> tableau = new ArrayList<ArrayList<String>>();

		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			// Now let's iterate over the columns of the current row
			cellIterator = row.cellIterator();
			ArrayList<String> rs = new ArrayList<String>();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				
				rs.add(readAsString(cell));
			}

			tableau.add(rs);

		}

		// Closing the workbook
		workbook.close();
		
		
		return tableau;
	}

	public String readAsString(Cell cell) throws ExcelErrorException {
		final String value;
		switch (cell.getCellTypeEnum()) {
		case STRING:
			value = cell.getStringCellValue().isEmpty() ? null : cell.getStringCellValue();
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				value = cell.getDateCellValue().toString();
			} else {
				return Double.toString(cell.getNumericCellValue());
			}
			break;
		case BOOLEAN:
			value = Boolean.toString(cell.getBooleanCellValue());
			break;
		case BLANK:
			value = null;
			break;
		case FORMULA:
			CreationHelper crateHelper = workbook.getCreationHelper();
			FormulaEvaluator evaluator = crateHelper.createFormulaEvaluator();
			final org.apache.poi.ss.usermodel.CellValue cellValue = evaluator.evaluate(cell);
			switch (cellValue.getCellTypeEnum()) {
			case STRING:
				value = cellValue.getStringValue();
				break;
			case NUMERIC:
				value = Double.toString(cellValue.getNumberValue());
				break;
			case BOOLEAN:
				value = Boolean.toString(cellValue.getBooleanValue());
				break;
			case ERROR:
				throw new ExcelErrorException("erreur dans la formule excel");
			default:
				throw new IllegalStateException();
			}
			break;
		case ERROR:
			throw new ExcelErrorException("erreur dans la cellule ");
		case _NONE:
		default:
			throw new IllegalStateException();
		}
		return value;
	}

}
