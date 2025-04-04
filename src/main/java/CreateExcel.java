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


        for (int i = 0; i < 100000; i++) {
            list.add(r.nextDouble(1.0, 100000.0));
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

        sheet.getRow(5).getCell(3).setCellValue(i);

        generateFrequency(i, valoresIntevals);

        setMidPoint();

        if (false) {


//        for ( i = 3; i <= infos.get("K") - 1; i++) {
//            rows.get(i).createCell(4).setCellFormula("F" + i);
//        }
//
//        for (i = 3; i <= infos.get("K") - 1 ; i++) {
//
//            String formula = "E";
//
//            formula += i + 1;
//
//            formula += "+" + c;
//
//            rows.get(i).createCell(5).setCellFormula(formula);
//        }


            var i2 = i - 1;

            FormulaEvaluator evaluator = w.getCreationHelper().createFormulaEvaluator();
            CellValue value = evaluator.evaluate(rows.get(i2).getCell(5));

            //Apenas para receber o valor da celula
            var x = value.getNumberValue();


            addClass(i2);

            i2 = i - 1;

            evaluator = w.getCreationHelper().createFormulaEvaluator();

            value = null;

            while (value == null) {
                value = evaluator.evaluate(rows.get(i2).getCell(5));
                i2--;
            }

            x = value.getNumberValue();

            i2++;

            while (x < infos.get("Xmax")) {

                if (sheet.getRow(i2) == null) {
                    rows.add(sheet.createRow(i2));
                }

                rows.get(i2).createCell(4).setCellFormula("F" + i2);
                rows.get(i2).createCell(5).setCellFormula("E" + (i2 + 1) + "+" + c);

                evaluator = w.getCreationHelper().createFormulaEvaluator();
                value = evaluator.evaluate(rows.get(i2).getCell(5));
                x = value.getNumberValue();

                var a = infos.get("K");
                infos.remove("K");
                infos.put("K", a + 1);

                i2++;

                rows.get(5).getCell(2).setCellValue(infos.get("K"));
            }


            String limitesSuperiores = "$F$3:$F$" + i2;


            for (i = 2; i < infos.get("K") + 2; i++) {

                double ls;
                double li;


                value = evaluator.evaluate(rows.get(i).getCell(4));
                li = value.getNumberValue();


                value = evaluator.evaluate(rows.get(i).getCell(5));
                ls = value.getNumberValue();

                //System.out.println(li + " - " + ls);

                NumberFormat frt = NumberFormat.getInstance(Locale.getDefault());

                // Formata o número de acordo com a localização do sistema
                //Pois quando ele envia a função, ele leva o numero separando
                //as casas decimais por "." e sistemas q estão usando "," ele nn funciona
                String liS = frt.format(li);
                String lsS = frt.format(ls);

                System.out.println(liS + " - " + lsS);

                rows.get(i).createCell(6).setCellFormula("COUNTIFS(" + valoresIntevals + ", " + "\"<=" + lsS + "\"," + valoresIntevals + ", \">=" + liS + "\")");
            }


            for (int i3 = 2, l = 3; i3 < infos.get("K") + 2; i3++, l++) {
                rows.get(i3).createCell(7).setCellFormula("(E" + l + "+ F" + l + ")/2");
            }

            for (int i3 = 2, l = 3; i3 < infos.get("K") + 2; i3++, l++) {
                rows.get(i3).createCell(8).setCellFormula("G" + l + "/$C$2");
            }

            for (int i3 = 2, l = 3; i3 < infos.get("K") + 2; i3++, l++) {
                rows.get(i3).createCell(9).setCellFormula("(G" + l + " * 100)" + "/$C$2");
            }

            rows.get(2).createCell(10).setCellFormula("G3");

            for (int i3 = 3, l = 4; i3 < infos.get("K") + 2; i3++, l++) {
                rows.get(i3).createCell(10).setCellFormula("G" + l + "+K" + (l - 1));
            }

            rows.get(2).createCell(11).setCellFormula("J3");

            for (int i3 = 3, l = 4; i3 < infos.get("K") + 2; i3++, l++) {
                rows.get(i3).createCell(11).setCellFormula("J" + l + "+L" + (l - 1));
            }

            rows.get(12).createCell(12).setCellFormula("G13");

            for (int i3 = 11, l = 12; i3 > 1; i3--, l--) {
                rows.get(i3).createCell(12).setCellFormula("G" + l + "+M" + (l + 1));
            }

            rows.get(12).createCell(13).setCellFormula("J13");

            for (int i3 = 11, l = 12; i3 > 1; i3--, l--) {
                rows.get(i3).createCell(13).setCellFormula("J" + l + "+N" + (l + 1));
            }


            var strings = new String[]{"Média", "Moda", "Mediana", "Variância", "Desvio Padrão"};

            for (int i3 = 8, l = 0; i3 < 13; i3++, l++) {


                if (sheet.getRow(i3) == null) {
                    sheet.createRow(i3);
                }
                if (sheet.getRow(i3).getCell(1) == null) {
                    rows.get(i3).createCell(1).setCellValue(strings[l]);
                } else {
                    sheet.getRow(i3).getCell(1).setCellValue(strings[l]);
                }

            }


            strings = new String[]{"AVERAGE(" + valoresIntevals + ")", "MODE(" + valoresIntevals + ")", "MEDIAN(" + valoresIntevals + ")", "VAR(" + valoresIntevals + ")", "STDEV(" + valoresIntevals + ")"};


            for (int i3 = 8, l = 0; i3 < 13; i3++, l++) {

                if (sheet.getRow(i3).getCell(2) == null) {
                    rows.get(i3).createCell(2).setCellFormula(strings[l]);
                } else {
                    rows.get(i3).getCell(2).setCellFormula(strings[l]);
                }

            }

            w.getCreationHelper().createFormulaEvaluator().evaluateAll();

        }

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
    public static void generateFrequency(int k, String intervals) {

        int i = 3, l = 0;


        while (l < k) {

            evaluator = w.getCreationHelper().createFormulaEvaluator();

            //Pega celula do limite inferior

            value = evaluator.evaluate(cell("E" + i));

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
            l++;

        }


    }

    public static void setMidPoint(){

        for(int i = 3; i < infos.get("K") + 3; i++){
            cell("H" + i).setCellFormula("(E" + i + " + F" + i + ")/2");
        }


    }

    //Função q retorne a celula por uma string => "C5" -> row.get(2).getCell(4)
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
