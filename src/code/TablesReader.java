package code;

/**
 *  Программа может переделать excel таблицу в формате .xlsx в .txt файл
 *  с выписанными в удобным для дальнейшей обработки формате, или в
 *  импортируемый в стандартные календари файл в формате .ics
 *
 *  Используемая версия JavaSE - 8
 *
 *  Класс Tables reader - считывает таблицу и записывает все данные
 *  в формате объектов класса Lesson в массив LessonData, который является
 *  private переменной и может быть считан или задан имеющимися getter'ами
 *  и setter'ами
 *
 *  На вход подается таблица, допускающая наличие одного расписания, в которой необходимо
 *  удалить все что выше первой строки с первым предметом и ниже последнего предмета во избежание
 *  ошибочного занесения в массив LessonData побочных данных ( очень склизский момент,
 *  алгоритмизация процесса подготовки таблицы к парсеру - почти бессмысленна, так как все таблицы
 *  разные и, почти всегда, неодинаковы для разных групп ).
 *
 *  В проект проекте используются 2 сторонних фреймворка:
 *   - Apache POI - фреймворк для работы с таблицами excel
 *   - Biweekly - фреймворк для работы с .ics файлами
 *
 *  Структура проекта:
 *  /code - папка с исполняемыми классами
 *  /res - ресурсы
 *  /res/cals - здесь создаются готовые календари
 *  /res/txts - здесь создаются готовые .txt файлы
 *  /res/xlsx - сюда нужно помещать исходники таблиц, названия которых содержат только номер группы
 *  файлы workbook_*.xlsx - промежуточный генерируемый результат обработки таблицы методом removeMerges(), в котором
 *  иногда проще увидеть артефакты таблицы
 *
 *  Для примера в папку с исходниками таблиц помещено готовое для парсера расписание 415ой группы ФМЭСИ на второй модуль 2015г.
 *
 *  <-------------------------------------------------------------------------------------------------------->
 */

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;

/**
 *  Главный исполняемый класс, содержащий main(), который управляет работой проекта
 */
public class TablesReader {

    /**
     * Статичное объявление переменных класса
     */
    private static XSSFRow row;
    private static XSSFCell cell;
    private static File file;
    private static FileInputStream fileInputStream;
    private static XSSFWorkbook workbook;
    private static XSSFSheet sheet;
    public static String[] days = {"Понедельник", "Вторник", "Среда",
            "Четверг", "Пятница", "Суббота"};

    /**
     * @param group - номер группы, расписание которой предстоит создать
     *              его стоит указывать перед компилляцией программы вручную
     *              и обязательно верно - таблицы ищутся и редактируются далее
     *              автоматически и полностью в зависимости от значения group
     */
    public static String group = "415";


    /**
     * @throws IOException - перехват ошибки при отсутствии нужного исходника в папке /res/xlsx
     * @throws ParseException - перехват ошибки при обработке таблицы
     */
    public static void main(String[] args) throws IOException, ParseException {
        // Определение текущей даты и создание готовых к использованию открытых переменных
        // для работы со временем
        DateData.setDate();

        // Подготока к парсеру файла workbook_*.xlsx
        removeMerges();

        // Считывание workbook_*.xlsx в объекты классов Lesson и LessonData
        readTable();


        CalendarCreator.create();
        TxtCreator.createTxt();
    }



    /**
     * @param s is "A7"-like input index
     * @return is convinient 2-element adress[]
     *         which has @row index as [0] element
     *         and @column index as [1] element
     */
    public static int[] getNumericIndex (String s){
        int rowNum;
        int[] adress = new int[2];
        rowNum = Integer.parseInt(s.substring(1));
        //row
        adress[0]=rowNum-1;
        //column
        s=s.toUpperCase();
        adress[1] = Character.codePointAt(s,0)-65;
        return adress;
    }

    private static char[] alphabet = {'A','B','C','D','E','F','G','H','I','G','K'};

    /**
     * Метод получает координаты ячейки
     * @param rowNum - номер строки
     * @param colNum - номер столбца
     * @return координаты ячейки в формате "А7", "В13" и т.д., более удобном для восприятия
     */
    public static String getStringIndex (int rowNum, int colNum){
        return alphabet[colNum+1]+Integer.toString(rowNum+1);
    }

    /**
     * Метод возвращает текстовое значение ячейки с координатами:
     * @param rowNum - номер строки
     * @param colNum - номер столбца
     */
    public static String getCellText(int rowNum, int colNum){
        row = sheet.getRow(rowNum);
        cell = row.getCell(colNum);
        if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
            return Integer.toString((int) cell.getNumericCellValue());
        }else return cell.getStringCellValue();
    }

    /**
     * Метод возвращает текстовое значение ячейки с координатами:
     * @param id - координата ячейки в представлении "A7, B13" и т.д.
     */
    public static String getCellText(String id){
        return getCellText(getNumericIndex(id)[0],getNumericIndex(id)[1]);
    }

    /**
     * Метод записывает в ячейку значение, переданное в
     * @param text
     *  с координатой
     * @param cellId
     */
    public static void setCellText(String cellId, String text){
        setCellText(getNumericIndex(cellId)[0], getNumericIndex(cellId)[1], text);
    }

    /**
     * Метод записывает в ячейку значение, переданное в
     * @param text
     *  с координатами
     * @param rowNum
     * @param cellNum
     */
    public static void setCellText(int rowNum, int cellNum, String text){
        row = sheet.getRow(rowNum);
        cell = row.getCell(cellNum);
        cell.setCellValue(text);
    }

    /**
     * Метод возвращает обект класса XFFSCell с координатами
     * @param rowNum
     * @param colNum
     */
    public static XSSFCell getCell(int rowNum, int colNum){
            System.out.println("row : "+ rowNum);
            System.out.println("column : " + colNum);
            row = sheet.getRow(rowNum);
            return row.getCell(colNum);
    }


    /**
     * Метод подгатавливает таблицу к парсеру
     *
     * В результате работы метода сохраненяется промежуточный результат обработки таблицы для дальнейшего
     * ее парсера workbook_*.xlsx
     * При дальнейшем преобраовании таблицы используется именно workbook_*.xlsx файл, а не исходник
     * При этом исходник таблицы не изменяется вовсе во избежание повреждения таблицы
     * в случае ошибки программы - перед началом обработки таблица копируется в память
     */
    private static void removeMerges(){
        try {
            openTable(group);

            String[] merged;
            while ( sheet.getNumMergedRegions()!=0){
                System.out.println(sheet.getMergedRegion(0).formatAsString());
                merged = sheet.getMergedRegion(0).formatAsString().split(":",2);

                if ( getNumericIndex(merged[0])[0] == getNumericIndex(merged[1])[0]){
                    setCellText(merged[1],getCellText(merged[0]));
                    System.out.println( " = " + getCellText(merged[0]));

                    for ( int k = 0 ; k<2; k++) {
                        String toCopyCellId = getStringIndex(getNumericIndex(merged[1])[0],
                                getNumericIndex(merged[1])[1] + k);
                        String text = getCellText(toCopyCellId);

                        String unmergedCellId = getStringIndex(getNumericIndex(merged[0])[0],
                                getNumericIndex(merged[0])[1] +k);

                        System.out.println(sheet.getMergedRegion(0).formatAsString()
                                + " : " + unmergedCellId + " -> " + text + " from " + toCopyCellId);

                        setCellText(unmergedCellId, text);
                    }
                } else
                    for ( int r = sheet.getMergedRegion(0).getFirstRow();
                             r<=sheet.getMergedRegion(0).getLastRow(); r++){
                        for ( int c = sheet.getMergedRegion(0).getFirstColumn();
                                c<=sheet.getMergedRegion(0).getLastColumn(); c++){
                            if(sheet.getMergedRegion(0).getFirstColumn()<2)
                            setCellText(r,c,getCellText(merged[0]));
                        }
                    }
                    sheet.removeMergedRegion(0);
                }
            saveTable("workbook_"+group);
        } catch (IOException e) {
            System.out.println("FILE_SHIT");
            e.printStackTrace();
        }
    }


    /**
     * Метод копирует в память таблицу с названием , переданным в
     * @param tableName
     * @throws IOException - перехват ошибки при открытии файла
     */
    private static void openTable(String tableName) throws IOException{
        file = new File("src/res/xlsx/"+tableName+".xlsx");
        fileInputStream = new FileInputStream(file);
        workbook = new XSSFWorkbook(fileInputStream);
        sheet = workbook.getSheetAt(0);
    }

    /**
     * Метод копирует таблицу из памяти в файл с именем
     * @param tableName
     * @throws IOException - перехват ошибки при записи файла
     */
    private static void saveTable(String tableName) throws IOException {
        FileOutputStream fileOut = new FileOutputStream("src/res/xlsx/" + tableName + ".xlsx");
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }


    // Массив lessonsData хранит объекты DayData на каждый день недели
    private static DayData[] lessonsData = new DayData[7];

    public static DayData getDayData(int n) {
        return lessonsData[n];
    }

    /**
     * Копирование готовой к парсеру workbook_*.xlsx таблицы в объекты Lesson
     * которые записываются в DayData по дням недели, а затем в LessonData
     * @throws IOException
     */
    private static void readTable() throws IOException {
        openTable("workbook_"+group);
        Lesson lesson1;
        String weekday,str, time1="",time2="";
        int rowCounter=-1;
        for ( int i=0 ; i<6; i++){
            // Создание объекта класса lessonsData
            lessonsData[i]= new DayData();
            weekday = getCellText(rowCounter+1,0);
            System.out.println(weekday.toUpperCase());
            do{
                rowCounter++;
                if (!(getCellText(rowCounter, 2).equals(""))) {
                    lesson1 = new Lesson();
                    System.out.println("Adding lesson to "+group+" on " +weekday);

                    //LESSON TIME
                    str = getCellText(rowCounter, 1).replaceAll(" ", "");
                    time1 = str.split("-", 2)[0];
                    time2 = str.split("-", 2)[1];
                    lesson1.setStartTime(time1);
                    lesson1.setEndTime(time2);
                    System.out.println(time1 + " - " + time2);

                    // LESSON NAME AND LECTURER NAME
                    str = getCellText(rowCounter,2);
                    if (str.contains("(")){
                        lesson1.setLecturer(str.substring(str.indexOf("(")+1,str.indexOf(")")));
                        lesson1.setLessonName(str.substring(0, str.indexOf("(")));}
                    else lesson1.setLessonName(str);

                    System.out.println(lesson1.getLessonName() + " - " + lesson1.getLecturer());

                    // LESSON TYPE
                    str = getCellText(rowCounter,3);
                    if (str.contains("л") || str.contains("Л")){
                        lesson1.setType("Лекция");
                        str = str.replace("л", "");
                        str = str.replace("Л", "");
                    }
                    if (str.contains("с") || str.contains("С")) {
                        lesson1.setType("Семинар");
                        str = str.replace("с", "");
                        str = str.replace("С", "");
                    }


                    // WEEKS
                    if ( str.contains("-")) {
                        lesson1.setStartWeek(Integer.parseInt(str.split("-", 2)[0]));
                        lesson1.setEndWeek(Integer.parseInt(str.split("-", 2)[1]));
                    } else {
                        lesson1.setStartWeek(Integer.parseInt(str));
                        lesson1.setEndWeek(lesson1.getStartWeek());
                    }


                    // CABINET
                    str = getCellText(rowCounter, 4);
                    if (str!= null && str.length()>0 && !str.contains("к")) str=str+" 3к";
                    lesson1.setCabinet(str);

                    lessonsData[i].addDATA(lesson1);
                }
            } while(weekday.equals(getCellText(rowCounter+1,0)));
        }
    }
}

/**
 * Данный класс хранит описание самих пар
 */
class Lesson{

    public Lesson(){}

    private String lessonName;
    private String cabinet;
    private String type;
    private int startWeek;
    private int endWeek;
    private String startTime;
    private String endTime;
    private String lecturer;

    public String getLecturer() {
        return lecturer;
    }

    public void setLecturer(String lecturer) {
        this.lecturer = lecturer;
    }

    public String getStartTime() {
        return startTime;
    }

    public void setStartTime(String startTime) {
        this.startTime = startTime;
    }

    public String getEndTime() {
        return endTime;
    }

    public void setEndTime(String endTime) {
        this.endTime = endTime;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public int getStartWeek() {
        return startWeek;
    }

    public void setStartWeek(int startWeek) {
        this.startWeek = startWeek;
    }

    public int getEndWeek() {
        return endWeek;
    }

    public void setEndWeek(int endWeek) {
        this.endWeek = endWeek;
    }

    public String getLessonName() {
        return lessonName;
    }

    public void setLessonName(String lessonName) {
        this.lessonName = lessonName;
    }

    public String getCabinet() {
        return cabinet;
    }

    public void setCabinet(String cabinet) {
        this.cabinet = cabinet;
    }
}

/**
 * Даный класс создает объекты, которые содержат все пары на каждый день недели
 */
class DayData extends ArrayList {
    private ArrayList<Lesson> DATA = new ArrayList<Lesson>();

    // Пустой конструктор подгатавливает объект класса к дальнейшему заполнению
    public DayData() {}

    public void addDATA(Lesson lesson) {
        this.DATA.add(lesson);
    }

    public ArrayList<Lesson> getDATA() {
        return DATA;
    }
}