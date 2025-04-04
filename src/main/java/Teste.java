public class Teste {

    public static void main(String[] args) {

        CreateExcel.cell("E3").setCellValue(100);


        var value = CreateExcel.cell("E3").getStringCellValue();

    System.out.println("Aqui: '" + value + "'");

    }

}
