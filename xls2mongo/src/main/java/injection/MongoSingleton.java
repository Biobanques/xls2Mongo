package injection;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Properties;

import com.mongodb.DB;
import com.mongodb.MongoClient;
import com.mongodb.MongoCredential;
import com.mongodb.ServerAddress;

/**
 * singleton pour ne pas avoir trop de "connexion ouverte" sur la base mongo et
 * obtenir une erreur de type : "too many files opened"<br>
 * TODO ameliorer via un singleton njecter avec spring?
 * 
 * @author nmalservet
 * 
 */
public class MongoSingleton {

	public MongoSingleton() {
	}

	public static DB db;

	public static DB getDb() {
		if (db == null) {
			MongoClient mongoClient;
			String dbConnection = "src/main/resources/dbConnection.properties";
			Properties dbConnectionProperties = new Properties();
			try {

				FileInputStream connectionProps = new FileInputStream(
						dbConnection);
				dbConnectionProperties.load(connectionProps);
				connectionProps.close();

			} catch (IOException e) {
				System.out
						.println("Impossible de charger le fichier de configuration");
			}
			try {
				MongoCredential credentials = MongoCredential
						.createMongoCRCredential(dbConnectionProperties
								.getProperty("login"), dbConnectionProperties
								.getProperty("dbConnection"),
								dbConnectionProperties.getProperty("password")
										.toCharArray());

				mongoClient = new MongoClient(new ServerAddress(
						dbConnectionProperties.getProperty("server"),
						Integer.parseInt(dbConnectionProperties
								.getProperty("port"))),
						Arrays.asList(credentials));

				db = mongoClient.getDB(dbConnectionProperties
						.getProperty("dbName"));
				// adnDb
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return db;
	}

}
