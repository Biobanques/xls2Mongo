package main;

import injection.MongoSingleton;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map.Entry;
import java.util.Properties;

import javax.swing.JOptionPane;

import com.mongodb.DB;
import com.mongodb.DBCollection;

import extraction.Extractor;

public class Run {

	public static void main(String[] args) {
		String excelFilePaths = "src/main/resources/filepaths.properties";
		Properties excelFilesProperties = new Properties();
		String dbConnection = "src/main/resources/filepaths.properties";
		Properties dbConnectionProperties = new Properties();
		try {
			FileInputStream excelProps = new FileInputStream(excelFilePaths);
			excelFilesProperties.load(excelProps);
			excelProps.close();
			FileInputStream connectionProps = new FileInputStream(dbConnection);
			excelFilesProperties.load(connectionProps);
			connectionProps.close();

		} catch (IOException e) {
			System.out
					.println("Impossible de charger le fichier de configuration");
		}

		try {
			DB db = MongoSingleton.getDb();
			DBCollection collection = db.getCollection("excelRow");
			collection.drop();
			// TODO : mettre la liste des fichiers à extraire dans un fichier de
			// conf, type XML
			HashMap<String, Integer> filesPath = new HashMap<String, Integer>();

			for (Entry<Object, Object> prop : excelFilesProperties.entrySet()) {
				String fileName = prop.getKey().toString();
				Integer nbSheets = Integer.valueOf(prop.getValue().toString());

				filesPath.put(fileName, nbSheets);
			}

			// boucle sur la liste de fichiers pour extraire les données
			for (Entry<String, Integer> filepath : filesPath.entrySet()) {
				Extractor.excelExtract(collection, filepath);
			}
			JOptionPane.showMessageDialog(null,
					"Extraction effectuée avec succès.");
		} catch (Exception e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(null,
					"Une erreur est survenue, l'exportation a échouée.");
		}

	}
}
