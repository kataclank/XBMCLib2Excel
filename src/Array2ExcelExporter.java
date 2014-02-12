/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * 
 * @author rafael.delarosa Este método recive un array de String[] y genera un
 *         .xsl y devuelve los bytes
 * 
 *         El primer elemento del array son las cabeceras de los elementos.
 */
public class Array2ExcelExporter {
	public static void doExcel(final Object[] datosExportar, final Short[] colores, final String excelResultFile)
			throws IOException {
		new Date();
		String pathArchivo = excelResultFile;
		pathArchivo = pathArchivo.replace("\\", "/");
		final HSSFWorkbook workbook = new HSSFWorkbook();
		final HSSFSheet hoja = workbook.createSheet("Listado_XBMC");
		final String titulos[] = (String[]) datosExportar[0];
		HSSFRow fila = hoja.createRow(0);
		final HSSFSheet sheet = workbook.getSheetAt(0);
		for (int i = 0; i < titulos.length; i++) {
			final String dat = titulos[i];
			final HSSFCell celda = fila.createCell(new Integer(i));
			final HSSFFont fuente = workbook.createFont();
			fuente.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			final HSSFCellStyle estilo = workbook.createCellStyle();
			estilo.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			if (colores != null) {
				estilo.setFillBackgroundColor(colores[i]);
				estilo.setFillForegroundColor(colores[i]);
				fuente.setColor(HSSFColor.BLACK.index);
			} else {
				estilo.setFillForegroundColor(HSSFColor.BLACK.index);
				estilo.setFillBackgroundColor(HSSFColor.BLACK.index);
				fuente.setColor(HSSFColor.WHITE.index);
			}
			estilo.setBorderTop(HSSFCellStyle.BORDER_THIN);
			estilo.setTopBorderColor(HSSFColor.BLACK.index);
			estilo.setBorderRight(HSSFCellStyle.BORDER_THIN);
			estilo.setRightBorderColor(HSSFColor.BLACK.index);
			estilo.setBorderBottom(HSSFCellStyle.BORDER_THIN);
			estilo.setBottomBorderColor(HSSFColor.BLACK.index);
			estilo.setBorderLeft(HSSFCellStyle.BORDER_THIN);
			estilo.setLeftBorderColor(HSSFColor.BLACK.index);
			estilo.setLocked(true);
			estilo.setFont(fuente);
			celda.setCellStyle(estilo);
			final HSSFRichTextString richText = new HSSFRichTextString(dat);
			// richText.applyFont(HSSFFont.ANSI_CHARSET);
			richText.applyFont(fuente);
			// richText.S
			celda.setCellValue(richText);
			sheet.autoSizeColumn((short) i);
		}

		final HSSFCellStyle estilo1 = workbook.createCellStyle();
		estilo1.setBorderTop(HSSFCellStyle.BORDER_THIN);
		estilo1.setTopBorderColor(HSSFColor.BLACK.index);
		estilo1.setBorderRight(HSSFCellStyle.BORDER_THIN);
		estilo1.setRightBorderColor(HSSFColor.BLACK.index);
		estilo1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		estilo1.setBottomBorderColor(HSSFColor.BLACK.index);
		estilo1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		estilo1.setLeftBorderColor(HSSFColor.BLACK.index);

		for (int i = 1; i < datosExportar.length; i++) {
			final String columnas[] = (String[]) datosExportar[i];
			if (columnas != null) {
				fila = hoja.createRow(i);
				for (int j = 0; j < columnas.length; j++) {
					sheet.autoSizeColumn((short) j);
					final String dat = columnas[j];
					if (org.apache.commons.lang.StringUtils.isNotBlank(dat)) {
						final HSSFCell cell = fila.createCell(new Integer(j));
						cell.setCellStyle(estilo1);
						final HSSFRichTextString richText = new HSSFRichTextString(dat);
						richText.applyFont(HSSFFont.ANSI_CHARSET);
						cell.setCellValue(richText);

					} else {
						final HSSFCell cell = fila.createCell(new Integer(j));
						cell.setCellStyle(estilo1);
						final HSSFRichTextString richText = new HSSFRichTextString("");
						richText.applyFont(HSSFFont.ANSI_CHARSET);
						cell.setCellValue(richText);
					}

				}
			}
		}
		final FileOutputStream fileOut = new FileOutputStream(pathArchivo);
		workbook.write(fileOut);
		workbook.getBytes();
		fileOut.close();
	}
}
