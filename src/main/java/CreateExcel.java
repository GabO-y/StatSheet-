import org.apache.commons.lang3.mutable.Mutable;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.*;

public class CreateExcel {

    static Workbook w = new XSSFWorkbook();
    static Sheet sheet = w.createSheet("Planilha");
    FormulaEvaluator evaluator = w.getCreationHelper().createFormulaEvaluator();
    CellValue valueCell;
    List<Double> values;
    String intervalValues;
    List<Row> rows = new ArrayList<>();
    Map<String, Double> infos;

    CellStyle style = w.createCellStyle();
    DataFormat format = w.createDataFormat();

    public CreateExcel(List<Double> values) {

        this.values = values;

        Collections.sort(values);

        setListValues(this.values);

        titles();

        infos = infos(values);

        classes();

        frequency(intervalValues);

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


    public static void main(String[] args) {

        List<Double> values = new ArrayList<>();

        values = List.of(
        19351.0,
        19352.0,
        19350.0,
        19353.0,
        19354.0,
        19355.0,
        19356.0,
        19357.0,
        19358.0,
        19359.0,
        19360.0,
        19361.0,
        19362.0,
        19363.0,
        19364.0,
        19365.0,
        19366.0,
        19367.0,
        19368.0,
        19369.0,
        19370.0,
        19371.0,
        19372.0,
        19373.0,
        19374.0,
        19375.0,
        19376.0,
        19377.0,
        19378.0,
        19379.0,
        19380.0,
        19381.0,
        19382.0,
        19383.0,
        19384.0
);

        values = new ArrayList<>(values);

//        var r = new Random();
//
//        for(int i = 0; i < 10000; i++){
//            values.add(r.nextDouble(1.0, 100000.0));
//        }

        CreateExcel tb = new CreateExcel(values);
    }
    public Map<String, Double> infos(List<Double> list) {

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

    public void setListValues(List<Double> list) {

        this.values = list;

        style.setDataFormat(format.getFormat("0.00"));

        for (int i = 2, l = 0; i < list.size() + 2; i++, l++) {
            cell("A" + i).setCellValue(values.get(l));
            cell("A" + i).setCellStyle(style);
        }

        this.intervalValues = "A2:A" + (values.size() + 1);
    }

    public void titles() {

        cell("A1").setCellValue("Valores");

        String[] information = new String[]{"N", "Xmax", "Xmin", "At", "K", "C"};

        for (int i = 2, l = 0; i <= 7; i++, l++) {
            cell("B" + i).setCellValue(information[l]);
        }

        cell("E2").setCellValue("Li");
        cell("F2").setCellValue("Ls");
        cell("G2").setCellValue("Fj");
        cell("H2").setCellValue("Xj");
        cell("I2").setCellValue("Fr");
        cell("J2").setCellValue("Fr%");
        cell("K2").setCellValue("F↓");
        cell("L2").setCellValue("F↓%");
        cell("M2").setCellValue("F↑");
        cell("N2").setCellValue("F↑%");

        information = new String[]{
                "COUNT(" + intervalValues + ")",
                "MAX(" + intervalValues + ")",
                "MIN(" + intervalValues + ")",
                "C3-C4",
                "1 + 3.3 * LOG10(C2)",
                "C5/C6"
        };

        for (int i = 2, l = 0; i <= 7; i++, l++) {
            cell("C" + i).setCellFormula(information[l]);
        }

        cell("D6").setCellFormula("ROUND(C6, 0)");
        cell("D7").setCellFormula("CEILING(C7, 1)");

        information = new String[] {"MEDIA", "MODA", "MEDIANA", "VARIÂNCIA", "DESVIO PADRÃO"};

        for(int i = 10, l = 0; i <= 14; i++, l++){
            cell("B" + i).setCellValue(information[l]);
        }

        information = new String[]{
                "AVERAGE(" + intervalValues + ")",
                "MODE(" + intervalValues + ")",
                "MEDIAN(" + intervalValues + ")",
                "VAR(" + intervalValues + ")",
                "STDEV(" + intervalValues + ")"
        };

        for(int i = 10, l = 0; i <= 14; i++, l++){
            cell("C" + i).setCellFormula(information[l]);
        }

        style.setDataFormat(format.getFormat("0.000"));

        for(int i = 3; i <= 14; i++){
            cell("C" + i).setCellStyle(style);
        }

    }
    public int addClass(int i2) {

        //Adiciona as duas primeiras linhas necessarias para construir as classes
        cell("E3").setCellFormula(infos.get("minL").toString());
        cell("F3").setCellFormula("E3+$D$7");

        evaluator = w.getCreationHelper().createFormulaEvaluator();
        valueCell = evaluator.evaluate(cell("F" + i2));

        var x = valueCell.getNumberValue();

        i2++;

        while (x < infos.get("Xmax")) {

//            if (sheet.getRow(i2) == null) {
//                rows.add(sheet.createRow(i2));
//                rows.get(i2).createCell(4);
//                rows.get(i2).createCell(5);
//            }
//            if (rows.get(i2).getCell(4) == null) {
//                rows.get(i2).createCell(4);
//            }
//            if (rows.get(i2).getCell(5) == null) {
//                rows.get(i2).createCell(5);
//            }

            cell("E" + i2).setCellFormula("F" + i2);
            cell("F" + i2).setCellFormula("E" + (i2 + 1) + " + " + "$D$7");

            evaluator = w.getCreationHelper().createFormulaEvaluator();
            valueCell = evaluator.evaluate(cell("F" + i2));

            //Apenas para receber o valor da celula
            x = valueCell.getNumberValue();

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

            valueCell = evaluator.evaluate(cell("E" + i));

            if (valueCell == null) {
                break;
            }

            var li = valueCell.getNumberValue();

            valueCell = evaluator.evaluate(cell("F" + i));

            var ls = valueCell.getNumberValue();

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

    public void classes(){


//        cell("D8").setCellFormula("LEN(D7)");
//
//        evaluator = w.getCreationHelper().createFormulaEvaluator();
//        valueCell = evaluator.evaluate(cell("D8"));
//
//        StringBuilder zeros = new StringBuilder("0");
//
//        for(int i = 1; i < valueCell.getNumberValue() - 1; i++){
//            zeros.append("0");
//        }
//
//        cell("D7").setCellFormula("CEILING(C7, 5" + zeros + ")");


        cell("E3").setCellValue(infos.get("minL"));
        cell("F3").setCellFormula("E3+$D$7");

        evaluator = w.getCreationHelper().createFormulaEvaluator();
        valueCell = evaluator.evaluate(cell("F3"));

        double x = valueCell.getNumberValue();

        int i = 4;

        while(infos.get("Xmax") > x){

            cell("E" + i).setCellFormula("F" + (i - 1));
            cell("F" + i).setCellFormula("E" + i + "+$D$7");

            evaluator = w.getCreationHelper().createFormulaEvaluator();
            valueCell = evaluator.evaluate(cell("F" + i));

            x = valueCell.getNumberValue();

            i++;

        }

    }
    public void frequency(String intervals) {

        for (int i = 3, l = 2; cell("F" + i, false) != null; i++, l++) {
            cell("G" + i).setCellFormula("COUNTIFS(" + intervals + "," + "\"<=\" & F" + i + " - 0.1, " + intervals + ", " + "\">=\" & E" + i + ")");
        }

        //=CONT.SES(A1:A100001; "<=" & F3; A1:A100001; ">=65")

    }

    public void setMidPoint() {

        for (int i = 3; sheet.getRow(i - 1).getCell(6) != null; i++) {
            cell("H" + i).setCellFormula("(E" + i + " + F" + i + ")/2");
        }


    }

    public void relativeFrequency() {

        for (int i = 3; sheet.getRow(i - 1).getCell(6) != null; i++) {
            cell("I" + i).setCellFormula("G" + i + "/$C$2");
        }

    }

    public void percentageFrequency() {

        for (int i = 3; sheet.getRow(i - 1).getCell(6) != null; i++) {
            cell("J" + i).setCellFormula("(G" + i + "*100)/$C$2");
        }

    }

    public void cumulativeBellow(String letterOfCell, String letterOfCumulate) {

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

    public void cumulativeAbove(String letterOfCell, String letterOfCumulate) {

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
    public Cell cell(String localization, Boolean createCell) {

        int letter = 0;
        int number = 0;

        if (!createCell) {

            var fields = localization.replaceAll(String.valueOf(localization.charAt(0)), localization.charAt(0) + ",").split(",");

            if (!(fields[0].charAt(0) >= 65 && fields[0].charAt(0) <= 90)) {
                System.out.println("Letter input incorrect");
                return sheet.getRow(0).getCell(0);
            }

            letter = (fields[0].charAt(0)) - 65;
            number = Integer.parseInt(fields[1]) - 1;


        } else {
            return cell(localization);
        }

        try {

            return sheet.getRow(number).getCell(letter);

        } catch (Exception _) {

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
