package com.conversaov1_0.conversaov1_0;

import com.google.common.hash.Hashing;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;

public class Conversao {

    //static Scanner teclado = new Scanner(System.in);

    public static void main(String[] args) throws IOException {

        //File diretorioInserido = new File(teclado.next());
        //System.out.println(diretorioInserido);

        //String diretorio = System.getProperty("user.home");
        //System.out.println(diretorio + "\\OneDrive - Iteris Consultoria e Software\\Área de Trabalho\\folder"); // Necessário fornecer o Deksktop certo
        //System.out.println(diretorio);

        File diretorioCompleto = new File("C:/Users/lucas.jacintho/OneDrive - Iteris Consultoria e Software/Área de Trabalho/folder"); // Necessário fornecer o Deksktop certo

        FileInputStream file = new FileInputStream(diretorioCompleto + "/Arquivo.xlsx"); // Abrindo o arquivo Excel
        Workbook planilhaOriginal = new XSSFWorkbook(file); // Importando o arquivo para o Java

        trabalharTabela(planilhaOriginal);

        salvarTabela(diretorioCompleto, planilhaOriginal);

        abrirNovaTabela();
    }

    private static void trabalharTabela(@NotNull Workbook planilhaOriginal)
    {
        Sheet sheet = planilhaOriginal.getSheetAt(0); // Pega a planilha 1 do arquivo Excel para trabalhar
        for (Row row : sheet) { //Para cada linha dentro da planilha faça:
            for (Cell cell : row) { //Para cada célula dentro da linha faça:

                //TRATAMENTO DA INFORMAÇÃO
                String valorOriginal = cell.getStringCellValue(); // Pega o valor da célula, sendo esse um valor númerico
                if (valorOriginal.length()>0) { // Verifica se de fato existe um valor a ser verificado
                    String valorConvertido = SHA256(valorOriginal); // Pega o valor String e envia pra o método Conversao

                    //MODIFICAÇÃO DA PLANILHA
                    Cell novaCell = row.createCell(1); //cria uma nova célula dentro da mesma linha do valor tratado
                    novaCell.setCellValue(valorConvertido.toUpperCase()); //Insere o valor obtido para dentro da célula recém criada
                }
            }
        }
    }

    private static @NotNull String SHA256(String valorConvertido) {
        return Hashing.sha256()
                .hashString(valorConvertido, StandardCharsets.UTF_8)
                .toString();
    }

    private static void salvarTabela(File diretorioCompleto, @NotNull Workbook planilhaOriginal) throws IOException {
//        try {
            String novoArquivo = diretorioCompleto + "/Convertido.xlsx"; // Aqui
            FileOutputStream arquivoFinal = new FileOutputStream(novoArquivo);//Aqui estamos criando o arquivo
            planilhaOriginal.write(arquivoFinal); //Aqui estamos escrevendo o arquivo
            planilhaOriginal.close();// Aqui fechamos a pasta
//        }
//        catch (Exception e){
//            System.out.println("Não foi possível executar a solicitação favor verifique se o arquivo está aberto, se sim, após fechar-lo tente novamente");
//            throw new IllegalAccessException();
//        }
    }

    private static void abrirNovaTabela() throws IOException {
        Desktop.getDesktop().open(new File( "C:/Users/lucas.jacintho/OneDrive - Iteris Consultoria e Software/Área de Trabalho/folder/Convertido.xlsx")); // Abrir o arquivo
    }

}