import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.AnyDocument;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.NumberFormat;
import java.util.*;

public class CreateExcel {

    static Workbook w = new XSSFWorkbook();
    static Sheet sheet = w.createSheet("Planilha");

    static FormulaEvaluator evaluator = w.getCreationHelper().createFormulaEvaluator();
    static CellValue value;
    static List<Row> rows = new ArrayList<>();

    static Map<String, Double> infos;

    public static void main(String[] args) {

        for (int i = 0; i <= 6; i++) {
            rows.add(sheet.createRow(i));
        }

        rows.get(1).createCell(1).setCellValue("n");
        rows.get(2).createCell(1).setCellValue("Xmax");
        rows.get(3).createCell(1).setCellValue("Xmin");
        rows.get(4).createCell(1).setCellValue("At");
        rows.get(5).createCell(1).setCellValue("K");
        rows.get(6).createCell(1).setCellValue("C");

        //var list = List.of(10.0, 34.9, 80.8, 22.4, 10.0, 34.9, 80.8, 22.4, 10.0, 34.9, 80.8, 22.4, 10.0, 34.9, 80.8, 22.4);

        List<Double> list = new ArrayList<>();


        var r = new Random();

        for (int i = 0; i < 10000; i++) {
            list.add(r.nextDouble(1.0, 10000000.0));
        }

        infos = infos(list);
        Collections.sort(list);

        CellStyle style = w.createCellStyle();
        DataFormat format = w.createDataFormat();
        style.setDataFormat(format.getFormat("0.00"));

        for (int i = 0, l = 1; i < list.size(); i++, l++) {

            if (l > 6) {

                rows.add(sheet.createRow(l));

            }
            rows.get(l).createCell(0).setCellValue(list.get(i));
            rows.get(l).getCell(0).setCellStyle(style);

        }

        //Intervalo dos valores
        String valoresIntevals = "A1:A" + (list.size() + 1);


        //Linha do N
        rows.get(1).createCell(2).setCellFormula("COUNT(" + valoresIntevals + ")");
        //Linha do Xmax
        rows.get(2).createCell(2).setCellFormula("MAX(" + valoresIntevals + ")");
        //Linha do Xmin
        rows.get(3).createCell(2).setCellFormula("MIN(" + valoresIntevals + ")");
        //Linha da Amplitude total
        rows.get(4).createCell(2).setCellFormula("C3-C4");

        //Linha do K e seu arredondamento
        rows.get(5).createCell(2).setCellFormula("1 + 3.3 * LOG10(C2)");
        rows.get(5).createCell(3).setCellFormula("ROUND(C6, 0)");

        //Linha do C e seu arredondamento
        rows.get(6).createCell(2).setCellFormula("C5/C6");
        rows.get(6).createCell(3).setCellFormula("CEILING(C7, 5)");

        String[] linha1 = new String[]{"Li", "Ls", "Fj", "Xj", "Fj", "Fj%", "F↓", "F↓%", "F↑", "F↑%"};

        for (int i = 4, l = 0; i < 14; i++, l++) {

            try {
                rows.get(1).createCell(i).setCellValue(linha1[l]);
            } catch (Exception e) {
                rows.get(1).getCell(i).setCellValue(linha1[l]);
            }

        }

        String c = "$D$7";

        int i = addClass(2);

        cell("D6").setCellValue(i);

        frequency(valoresIntevals);

        setMidPoint();

        relativeFrequency();

        percentageFrequency();

        cumulativeBellow("K", "G");
        cumulativeBellow("L", "J");

        cumulativeAbove("M", "G");
        cumulativeAbove("N", "J");
        
        try (FileOutputStream fileOutput = new FileOutputStream("ArquivoCriado.xlsx")) {

            w.write(fileOutput);

            w.close();

            System.out.println("Arquivo criado com sucesso");

        } catch (Exception e) {
            System.out.println("Erro ao criar arquivo\n" + e.getMessage());
        }


    }

    public static Map<String, Double> infos(List<Double> list) {

        Map<String, Double> infos = new HashMap<>();

        double min = Double.MAX_VALUE;
        double max = Double.MIN_VALUE;

        for (double i : list) {
            if (i <= min) {
                min = i;
            }
            if (i >= max) {
                max = i;
            }
        }


        infos.put("Xmin", min);
        infos.put("Xmax", max);


        infos.put("At", (max - min));

        double k = 1 + 3.33 * Math.log10(list.size());

        k = Math.round(k);

        infos.put("K", k);

        long c = Math.round(infos.get("At") / infos.get("K"));

        while (c % 5 != 0) {
            c++;
        }

        long minL = Math.round(min);

        while (minL % 5 != 0) {
            minL--;
        }

        infos.put("C", (double) c);

        infos.put("minL", (double) minL);

        while (infos.get("minL") > min) {

            var a = infos.get("minL");
            infos.remove("minL");

            while (a % 5 != 0) {
                a--;
            }

            infos.put("minL", a);

        }

        return infos;

    }

    public static int addClass(int i2) {

        //Adiciona as duas primeiras linhas necessarias para construir as classes
        sheet.getRow(2).createCell(4).setCellFormula(infos.get("minL").toString());
        sheet.getRow(2).createCell(5).setCellFormula("E3+$D$7");

        evaluator = w.getCreationHelper().createFormulaEvaluator();
        value = evaluator.evaluate(sheet.getRow(i2).getCell(5));

        var x = value.getNumberValue();

        i2++;

        while (x < infos.get("Xmax")) {

            if (sheet.getRow(i2) == null) {
                rows.add(sheet.createRow(i2));
                rows.get(i2).createCell(4);
                rows.get(i2).createCell(5);
            }
            if (rows.get(i2).getCell(4) == null) {
                rows.get(i2).createCell(4);
            }
            if (rows.get(i2).getCell(5) == null) {
                rows.get(i2).createCell(5);
            }

            sheet.getRow(i2).getCell(4).setCellFormula("F" + i2);
            sheet.getRow(i2).getCell(5).setCellFormula("E" + (i2 + 1) + " + " + "$D$7");

            evaluator = w.getCreationHelper().createFormulaEvaluator();
            value = evaluator.evaluate(sheet.getRow(i2).getCell(5));

            //Apenas para receber o valor da celula
            x = value.getNumberValue();

            i2++;

        }
        return i2 - 2;

    }

    //Antiga forma de pegar frequencia
    /*
    public static void generateFrequency(String intervals) {

        int i = 3;

        while (true) {

            evaluator = w.getCreationHelper().createFormulaEvaluator();

            //Pega celula do limite inferior

            value = evaluator.evaluate(cell("E" + i));

            if (value == null) {
                break;
            }

            var li = value.getNumberValue();

            value = evaluator.evaluate(cell("F" + i));

            var ls = value.getNumberValue();

            NumberFormat frt = NumberFormat.getInstance(Locale.getDefault());

            // Formata o número de acordo com a localização do sistema
            //Pois quando ele envia a função, ele leva o numero separando
            //as casas decimais por "." e sistemas q estão usando "," ele nn funciona

            String liS = frt.format(li);
            String lsS = frt.format(ls);


            cell("G" + i).setCellFormula("COUNTIFS(" + intervals + ", " + "\"<=" + lsS + "\"," + intervals + ", \">=" + liS + "\")");

            i++;
        }

    }
    */
    public static void frequency(String intervals){

        for(int i = 3, l = 2; cell("F"+i,false) != null; i++, l++){
            cell("G" + i).setCellFormula("COUNTIFS(" + intervals + "," + "\"<=\" & F" + i + ", " + intervals + ", " + "\">=\" & E" + i + ")");
        }

        //=CONT.SES(A1:A100001; "<=" & F3; A1:A100001; ">=65")

    }
    public static void setMidPoint() {

        for (int i = 3; sheet.getRow(i - 1).getCell(6) != null; i++) {
            cell("H" + i).setCellFormula("(E" + i + " + F" + i + ")/2");
        }


    }

    public static void relativeFrequency() {

        for (int i = 3; sheet.getRow(i - 1).getCell(6) != null; i++) {
            cell("I" + i).setCellFormula("G" + i + "/$C$2");
        }

    }

    public static void percentageFrequency() {

        for (int i = 3; sheet.getRow(i - 1).getCell(6) != null; i++) {
            cell("J" + i).setCellFormula("(G" + i + "*100)/$C$2");
        }

    }

    public static void cumulativeBellow(String letterOfCell, String letterOfCumulate) {

        if (letterOfCumulate.length() != 1 || letterOfCell.length() != 1) {
            System.out.println("Letter provided incorrect");
            return;
        }


        for (int i = 3; sheet.getRow(i - 1).getCell(6) != null; i++) {

            if (i == 3) {
                cell(letterOfCell + i).setCellFormula(letterOfCumulate + i);
                continue;
            }

            cell(letterOfCell + i).setCellFormula(letterOfCumulate + i + "+" + letterOfCell + (i - 1));

        }

    }

    public static void cumulativeAbove(String letterOfCell, String letterOfCumulate) {

        if (letterOfCumulate.length() != 1 || letterOfCell.length() != 1) {
            System.out.println("Letter provided incorrect");
            return;
        }

        int tam = 3;

        while (sheet.getRow(tam - 1).getCell(6) != null) {
            tam++;
        }

        tam--;

        for (int i = tam; sheet.getRow(i - 2).getCell(6) != null; i--) {

            if (i == tam) {
                cell(letterOfCell + i).setCellFormula(letterOfCumulate + i);
                continue;
            }

            cell(letterOfCell + i).setCellFormula(letterOfCumulate + i + "+" + letterOfCell + (i + 1));

        }

    }


    //Função q retorne a celula por uma string => "C5" -> row.get(2).getCell(4)
    public static Cell cell(String localization, Boolean createCell) {

        int letter = 0;
        int number = 0;

        if(!createCell){

        var fields = localization.replaceAll(String.valueOf(localization.charAt(0)), localization.charAt(0) + ",").split(",");

        if (!(fields[0].charAt(0) >= 65 && fields[0].charAt(0) <= 90)) {
            System.out.println("Letter input incorrect");
            return sheet.getRow(0).getCell(0);
        }

        letter = (fields[0].charAt(0)) - 65;
        number = Integer.parseInt(fields[1]) - 1;


        }else{
            return cell(localization);
        }

        try{

        return sheet.getRow(number).getCell(letter);

        }catch(Exception _){

            return null;

        }

    }

    public static Cell cell(String localization) {

        var fields = localization.replaceAll(String.valueOf(localization.charAt(0)), localization.charAt(0) + ",").split(",");

        if (!(fields[0].charAt(0) >= 65 && fields[0].charAt(0) <= 90)) {
            System.out.println("Letter input incorrect");
            return sheet.getRow(0).getCell(0);
        }

        int letter = (fields[0].charAt(0)) - 65;
        int number = Integer.parseInt(fields[1]) - 1;


        if (sheet.getRow(number) == null) {
            sheet.createRow(number);
        }

        if (sheet.getRow(number).getCell(letter) == null) {
            sheet.getRow(number).createCell(letter);
        }

        return sheet.getRow(number).getCell(letter);

    }


}
