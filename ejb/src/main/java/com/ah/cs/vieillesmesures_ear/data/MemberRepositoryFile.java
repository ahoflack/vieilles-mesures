package com.ah.cs.vieillesmesures_ear.data;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import javax.enterprise.context.ApplicationScoped;

import com.ah.cs.vieillesmesures_ear.model.MemberFile;

@ApplicationScoped
public class MemberRepositoryFile {

	private final String cheminsfichiersxcel = "C:/Users/alexandre/ressourcesViellesMesures/donneesinitiales";
	
	private static List<MemberFile> liste = new  ArrayList<>();;
	
    public List<MemberFile> findAllOrderedByName() {
    	
    	try {
    	
    		liste = new ArrayList<>();
    		
    	Files.newDirectoryStream(Paths.get(cheminsfichiersxcel), "*.xlsx")
        .forEach(MemberRepositoryFile::ajouterFichier);
    	
    	} catch(IOException e) {
    		e.printStackTrace();
    	}
    	return liste;
    	
    	
    }
    
    private static void ajouterFichier(Path path) {
    	liste.add(new MemberFile(path));
    }
    
}
