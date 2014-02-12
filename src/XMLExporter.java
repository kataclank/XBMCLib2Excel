import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.lang.StringUtils;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

public class XMLExporter {

	public static String contenedorHD = "mkv";

	public static void main(final String[] args) {
		System.out.println("Introduce el fichero y la ruta del xml del XBMC, Ej: /home/luis/libreria.xml:");
		final Scanner sc = new Scanner(System.in);
		final String cadena = sc.nextLine();

		File file = null;
		try {
			file = new File(cadena);
		} catch (final Exception e) {
			e.printStackTrace();
		}
		System.out.println("Introduce el fichero y la ruta del excel resultado, Ej: /home/luis/listadoDestino.xls:");
		final String destino = sc.nextLine();
		sc.close();
		procesarXMLbytes(file, destino);
	}

	public static void procesarXMLbytes(final File fis, final String destino) {
		final DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
		DocumentBuilder db;
		final String[] cabecera = { "id", "resolucion", "hd", "dual", "titulo", "ruta" };
		try {
			db = dbf.newDocumentBuilder();
			final org.w3c.dom.Document doc = db.parse(fis);
			doc.getDocumentElement().normalize();
			final NodeList nodeLst = doc.getElementsByTagName("movie");
			new ArrayList<String>();
			final Integer tam = nodeLst.getLength();
			final Object[] excelRows = new Object[tam + 1];
			excelRows[0] = cabecera;
			System.out.print("Procesando el xml");
			for (int s = 0; s < tam; s++) {
				final Node fstNode = nodeLst.item(s);
				if (fstNode.getNodeType() == Node.ELEMENT_NODE) {
					System.out.print(".");
					final String movieTitle = getInfoNodo("title", fstNode);
					final String movieFile = getInfoNodo("filenameandpath", fstNode);
					final String id = getInfoNodo("id", fstNode);
					final String height = getInfoNodo("height", fstNode);
					final String width = getInfoNodo("width", fstNode);
					final Boolean dual = isDual("streamdetails", fstNode);
					final String dualStr = dual ? "[DUAL]" : "";
					String esHD = "";
					if (width.equals("0")) {
						if (StringUtils.endsWithIgnoreCase(movieFile, contenedorHD)) {
							esHD = "[¿HD?]";
						}
					} else {
						esHD = Long.parseLong(width) > 1100 ? "[HD]" : "";
					}
					final String[] row = { id, width + "x" + height, esHD, dualStr, movieTitle, movieFile };
					excelRows[s + 1] = row;
				}
			}
			System.out.println("");
			System.out.println("Finalizado el proceso de lectura");
			System.out.println("***************************************************");
			System.out.println("Generando el excel....");
			Array2ExcelExporter.doExcel(excelRows, null, destino);
			System.out.println("Excel generado en:" + destino);
		} catch (final ParserConfigurationException e) {
			e.printStackTrace();
		} catch (final SAXException e) {
			e.printStackTrace();
		} catch (final IOException e) {
			e.printStackTrace();
		}
	}

	public static Boolean isDual(final String nodeName, final Node fstNode) {
		Integer cont = 0;
		try {
			final Element fstElmnt = (Element) fstNode;
			final NodeList node = fstElmnt.getElementsByTagName(nodeName);
			final Element element = (Element) node.item(0);
			final NodeList contenido = element.getChildNodes();
			for (int i = 0; i < contenido.getLength(); i++) {
				final String nodeAudioName = contenido.item(i).getNodeName();
				if (StringUtils.equalsIgnoreCase(nodeAudioName, "audio")) {
					cont++;
				}
			}
		} catch (final Exception e) {
		}
		if (cont > 1) {
			return true;
		} else {
			return false;
		}
	}

	public static String getInfoNodo(final String nodeName, final Node fstNode) {
		String nodeStr = "";
		try {
			final Element fstElmnt = (Element) fstNode;
			final NodeList node = fstElmnt.getElementsByTagName(nodeName);
			final Element element = (Element) node.item(0);
			final NodeList contenido = element.getChildNodes();
			nodeStr = contenido.item(0).getNodeValue();
		} catch (final Exception e) {
			nodeStr = "0";
		}
		return nodeStr;
	}

}
