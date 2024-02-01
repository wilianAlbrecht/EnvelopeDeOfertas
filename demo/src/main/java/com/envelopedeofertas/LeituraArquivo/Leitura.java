package com.envelopedeofertas.LeituraArquivo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class Leitura {

	private String diretorioNomes;
	private String diretoriosModelos;
	private String diretorioModelosSaida;

	public void lerArquivos() {

		diretorioNomes = "C:\\Users\\wilian.albrecht\\Desktop\\envelopes\\nomes.txt";
		diretoriosModelos = "C:\\Users\\wilian.albrecht\\Desktop\\envelopes\\Modelos\\MODELO RECIBO MENSAL.docx";
		diretorioModelosSaida = "C:\\Users\\wilian.albrecht\\Desktop\\envelopes\\Envelopes prontos";
		
		try {

			// salva os nomes em uma lista
			List<String> nomes = Files.readAllLines(Paths.get(diretorioNomes));  //Olívio Leonhardt

			for (String nome : nomes) {

				 try (FileInputStream fis = new FileInputStream(diretoriosModelos);
			             XWPFDocument document = new XWPFDocument(fis)) {

			            for (XWPFParagraph paragraph : document.getParagraphs()) {
			                for (XWPFRun run : paragraph.getRuns()) {
			                    String text = run.getText(0);
			                    if (text != null && text.contains("nome")) {
			                        String newText = text.replaceAll("\\b" + "nome" + "\\b", nome);
			                        run.setText(newText, 0);
			                    }
			                }
			            }

			            FileOutputStream fos = new FileOutputStream(diretorioModelosSaida + "\\recibo Mensal " + nome + ".docx");
			            document.write(fos);
			            fos.close();

			        } catch (IOException e) {
			            System.err.println("Error reading or writing files: " + e.getMessage());
			        } 
			}

		} catch (IOException e) {
			System.err.println("Error reading or writing files: " + e.getMessage());
		}

	}

}
