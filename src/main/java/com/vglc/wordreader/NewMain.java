/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.vglc.wordreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 *
 * @author ebitware201703
 */
public class NewMain {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {   
        try {            
            FileInputStream fis = new FileInputStream(new File("C:\\Users\\e-bitware\\Documents\\PERL\\proy\\asdasd.docx"));
            XWPFDocument document = new XWPFDocument(fis);  
            String wordToFind = "mientras";
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (XWPFParagraph para : paragraphs) {                                
                if(para.getText().contains(wordToFind)) {
//                    la palabra se encuentra en este parrafo.
                    List<XWPFRun> oldRuns = para.getRuns();                                      
//                    obtener los n runs que tenga la palabra que estamos buscando
                    for(int mainCounter = 0; mainCounter < oldRuns.size(); mainCounter ++) {
                        XWPFRun mainRun = oldRuns.get(mainCounter);
                        String runText =  mainRun.getText(0);
                        if(runText != null && (runText.length() < wordToFind.length()) && runText.equals(wordToFind.substring(0, runText.length()))) {
//                            se crea funcionalidad para que sea el run inmediato
                            if((mainCounter + 1) < oldRuns.size()) {
                                XWPFRun secondRun = oldRuns.get(mainCounter + 1);
                                String secondRunText =  secondRun.getText(0);
                                if((secondRunText.length() < wordToFind.length()) && secondRunText.equals(wordToFind.substring(secondRunText.length(), wordToFind.length()))) {                                                                       
                                    mainRun.setText("nuevotexto", 0);
                                    mainRun.setItalic(false);
                                    mainRun.setUnderline(UnderlinePatterns.SINGLE);
                                    System.out.println("Palabra encontrada con dos diferentes estilos. Pimer estilo: " + runText + " - cursiva: " + mainRun.isItalic() +" - Subrayada: " +mainRun.getUnderline());
                                    System.out.println("Segundo estilo: " + secondRunText + " - cursiva: " + secondRun.isItalic() +" - Subrayada: " + secondRun.getUnderline());                                    
                                    para.removeRun( mainCounter + 1);
                                }
                            }
                        }else if(runText.equals(wordToFind)) {
                            mainRun.setText("nuevotexto", 0);
                            mainRun.setItalic(false);
                            mainRun.setUnderline(UnderlinePatterns.SINGLE);
                        }
                    }
                }                                     
            }            
            document.write(new FileOutputStream("C:\\Users\\e-bitware\\Documents\\PERL\\proy\\asdasd2.docx"));            
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }      
    }              
}
