package com.envelopedeofertas.LeituraArquivo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;

public class Leitura {

	private String diretorioNomes;
	private String diretoriosModelos;
	private String diretorioModelosSaida;
	private String diretorioAuxiliar;

	public void lerArquivos() {

		// diretorioNomes = "C:\\Users\\wilian.albrecht\\Desktop\\envelopes\\nomes.txt";
		diretoriosModelos = "C:\\Users\\wilian.albrecht\\Desktop\\recibo";
		diretorioModelosSaida = "C:\\Users\\wilian.albrecht\\Desktop\\recibos_prontos";
		diretorioAuxiliar = "C:\\Users\\wilian.albrecht\\Desktop\\recibo\\auxiliar";

		try {

			String[] arquivos = new File(diretoriosModelos).list();

			for (String arquivo : arquivos) {

				// File file = new File(diretoriosModelos + "\\" + arquivo);
				if (arquivo.endsWith(".odt") || arquivo.endsWith(".ott")) {
					converterODFParaDOTX(arquivo);
				} else if (arquivo.endsWith(".dotx") || arquivo.endsWith(".docx")) {
					editarDOTXFile(diretoriosModelos + "\\" + arquivo, arquivo);
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// public void editarArquivo(String fileName) {

	// // diretoriosModelos + "\\" + fileName

	// if (fileName.endsWith(".dotx")) {
	// editarDOTXFile(fileName);
	// }

	// }

	public void editarDOTXFile(String filePath, String fileName) {

		try (FileInputStream fis = new FileInputStream(filePath);
				XWPFDocument documento = new XWPFDocument(fis)) {

			for (XWPFParagraph p : documento.getParagraphs()) {
				for (XWPFRun run : p.getRuns()) {
					String texto = run.getText(0);
					if (texto != null && (texto.contains("2025") || texto.contains("Decimo"))) {
						texto = texto.replace("Decimo", "Décimo");
						texto = texto.replace("2025", "2026");
						run.setText(texto, 0);
					}
				}
			}

			try (FileOutputStream fos = new FileOutputStream(diretorioModelosSaida + "\\" + fileName)) {
				documento.write(fos);
				System.out.println("Palavra substituída com sucesso! Arquivo salvo");
			}

		} catch (

		IOException e) {
			System.err.println(fileName);
			e.printStackTrace();
		}
	}

	public void converterODFParaDOTX(String fileName) {

		String officeHome = "C:\\Program Files\\LibreOffice";

		final LocalOfficeManager officeManager = LocalOfficeManager.builder()
				.officeHome(officeHome) // opcional, só se o LibreOffice não estiver no PATH
				.install().build();

		try {

			officeManager.start();

			File inputFile = new File(diretoriosModelos + "\\" + fileName);

			String newFileName = fileName.replace(".odt", ".dotx");

			File outputFile = new File(diretorioAuxiliar + "\\" + newFileName);

			LocalConverter
					.builder()
					.build()
					.convert(inputFile)
					.to(outputFile)
					.as(DefaultDocumentFormatRegistry.DOTX) // saída como DOTX
					.execute();

			System.out.println("Conversão concluída!");

			officeManager.stop();

			editarDOTXFile(diretorioAuxiliar + "\\" + newFileName, newFileName);

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// public void editarODTFile(String fileName) {

	// try {
	// TextDocument document = TextDocument.loadDocument(diretoriosModelos + "\\" +
	// fileName);

	// TextNavigation search = new TextNavigation("2025", document);

	// while (search.hasNext()) {
	// TextSelection item = (TextSelection) search.nextSelection();
	// item.replaceWith("2026");
	// }

	// document.save(diretorioModelosSaida + "\\" + fileName);
	// System.out.println("Palavra substituída com sucesso! Arquivo salvo");

	// } catch (Exception e) {
	// System.err.println(fileName);
	// e.printStackTrace();

	// }

	// // Salvar o documento modificado
	// try (FileOutputStream out = new FileOutputStream(diretorioModelosSaida +
	// fileName)) {
	// document.write(out);
	// }

	// } catch (IOException e) {
	// e.printStackTrace();
	// }

}

// try {

// // salva os nomes em uma lista
// List<String> nomes = Files.readAllLines(Paths.get(diretorioNomes)); //Olívio
// Leonhardt

// for (String nome : nomes) {

// try (FileInputStream fis = new FileInputStream(diretoriosModelos);
// XWPFDocument document = new XWPFDocument(fis)) {

// for (XWPFParagraph paragraph : document.getParagraphs()) {
// for (XWPFRun run : paragraph.getRuns()) {
// String text = run.getText(0);
// if (text != null && text.contains("nome")) {
// String newText = text.replaceAll("\\b" + "nome" + "\\b", nome);
// run.setText(newText, 0);
// }
// }
// }

// FileOutputStream fos = new FileOutputStream(diretorioModelosSaida + "\\recibo
// Mensal " + nome + ".docx");
// document.write(fos);
// fos.close();

// } catch (IOException e) {
// System.err.println("Error reading or writing files: " + e.getMessage());
// }
// }

// } catch (IOException e) {
// System.err.println("Error reading or writing files: " + e.getMessage());
// }