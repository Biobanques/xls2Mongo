package extraction;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map.Entry;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.mongodb.BasicDBObject;
import com.mongodb.DBCollection;

public class Extractor {
	/***
	 * Format de date pour convertir les dates excel
	 */
	private static final String DATE_FORMAT = "ddMMyyyy";
	private static final Logger LOGGER = Logger.getLogger(Extractor.class);

	public static void excelExtract(DBCollection collection,
			Entry<String, Integer> filepath) {
		try {

			String path = filepath.getKey();
			int nbSheets = filepath.getValue();
			// ouverture du flux
			File file = new File(path);
			FileInputStream fis = new FileInputStream(path);
			// ouverture du fichier excel
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			// fermeture du flux
			fis.close();
			// parcours des feuilles du fichier, basée sur le nombre de feuilles
			// fourni
			for (int i = 0; i < nbSheets; i++) {
				XSSFSheet activeSheet = wb.getSheetAt(i);
				Iterator<Row> rowIter = activeSheet.rowIterator();
				// passage du curseur à la première ligne, considérée comme la
				// ligne d'entete
				XSSFRow enteteRow = (XSSFRow) rowIter.next();
				Iterator<Cell> cellIter = enteteRow.cellIterator();
				int nbCol = 0;
				// récupération du nombre de colonnes à exploiter
				while (cellIter.hasNext()) {
					nbCol++;
					cellIter.next();
				}
				XSSFRow activeRow = null;
				/*
				 * parcours des lignes de la feuille. Si la première cellule
				 * d'une ligne est vide, on arrete le parcours.
				 */
				while (rowIter.hasNext()) {
					activeRow = (XSSFRow) rowIter.next();
					if (activeRow.getCell(0) == null
							|| activeRow.getCell(0).toString().isEmpty()) {

						break;
					}
					BasicDBObject obj = new BasicDBObject();
					obj.append("folderSource", new String(file.getParent()));
					obj.append("fileName", new String(file.getName()));
					obj.append("sheetName",
							new String(activeSheet.getSheetName()));
					obj.append("rowNumber", activeRow.getRowNum() + 1);
					ArrayList<BasicDBObject> attributes = new ArrayList<BasicDBObject>();
					for (int numCol = 0; numCol < nbCol; numCol++) {
						BasicDBObject attribute = new BasicDBObject();
						attribute.append("key",
								getCellData(enteteRow.getCell(numCol)));
						attribute.append("value",
								getCellData(activeRow.getCell(numCol)));
						attributes.add(attribute);
					}
					obj.append("attributes", attributes);

					collection.insert(obj);
				}
			}
		} catch (Exception e) {
			LOGGER.error(e.getStackTrace().toString());
			e.printStackTrace();
		}

	}

	protected static String getCellData(XSSFCell myCell) throws Exception {
		String cellData;
		if (myCell == null) {
			return null;

		} else {
			switch (myCell.getCellType()) {
			case XSSFCell.CELL_TYPE_STRING:
				cellData = myCell.getStringCellValue();
				cellData = cellData.replaceAll(
						System.getProperty("line.separator"), " ");
				break;
			case XSSFCell.CELL_TYPE_NUMERIC:
				cellData = getNumericValue(myCell);
				break;
			case XSSFCell.CELL_TYPE_FORMULA:
				switch (myCell.getCachedFormulaResultType()) {
				case XSSFCell.CELL_TYPE_STRING:
					cellData = myCell.getStringCellValue();
					cellData = cellData.replaceAll(
							System.getProperty("line.separator"), " ");
					break;
				case XSSFCell.CELL_TYPE_NUMERIC:
					cellData = getNumericValue(myCell);
					break;
				default:
					cellData = null;
				}
				break;
			default:
				cellData = null;
			}
		}
		return cellData;
	}

	protected static String getNumericValue(XSSFCell myCell) throws Exception {
		if (myCell == null) {
			return null;
		}
		String cellData = "";

		if (HSSFDateUtil.isCellDateFormatted(myCell)) {
			SimpleDateFormat sdf = new SimpleDateFormat(DATE_FORMAT);
			cellData = sdf.format(myCell.getDateCellValue());
		} else {
			DataFormatter df = new DataFormatter();
			cellData = df.formatCellValue(myCell);
		}
		return cellData;
	}

}
