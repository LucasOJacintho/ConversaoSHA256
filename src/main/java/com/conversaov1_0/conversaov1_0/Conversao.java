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

    public static void main(String[] args) throws IOException {

        File diretorioCompleto = new File("C:/Converter"); // Necessário fornecer o Deksktop certo

        Workbook planilhaOriginal;

        try  {
            FileInputStream planilha = new FileInputStream(diretorioCompleto + "/Arquivo.xlsx"); // Abrindo o arquivo Excel
            planilhaOriginal = new XSSFWorkbook(planilha); // Importando o arquivo para o Java
        }
        catch(Exception e){
            System.out.println("O Arquivo não foi encontrado\n Qualquer dúvida sobre a utilização desse software, recomendamos a leitura de seu Readme");
            throw new IOException();
        }

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
                String valorOriginal = cell.getStringCellValue(); // Pega o valor da célula, sendo esse uma string
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
            String novoArquivo = diretorioCompleto + "/Convertido.xlsx"; // Aqui

            FileOutputStream arquivoFinal = new FileOutputStream(novoArquivo);//Aqui estamos criando o arquivo
            planilhaOriginal.write(arquivoFinal); //Aqui estamos escrevendo o arquivo
            planilhaOriginal.close();// Aqui fechamos a pasta
    }

    private static void abrirNovaTabela() throws IOException {
        Desktop.getDesktop().open(new File( "C:/Converter/Convertido.xlsx")); // Abrir o arquivo
    }

}