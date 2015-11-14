package code;

import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.ParseException;

public class TxtCreator {
    /**
     * Format of lessons_.txt file:
     *
     * weekDay(string)
     * lessonName(string)
     * lecturer
     * startTime
     * endTime
     * startWeek(int)
     * endWeek(int)
     * lesson type("л" or "с")
     * cabinet(string)
     * [empty line]
     * "endDay"
     *
     */

    public static void createTxt() throws IOException, ParseException {

        File file = new File("src/res/txts/schedule_"+TablesReader.group+".txt");
        PrintWriter out = new PrintWriter(file);
            System.out.println("Writing "+TablesReader.group);

            for( int i = 0; i<6; i++){
                out.println(TablesReader.days[i]);
                for (Lesson lesson : TablesReader.getDayData(i).getDATA()) {
                    out.println(lesson.getLessonName());
                    out.println(lesson.getLecturer());
                    out.println(lesson.getStartTime());
                    out.println(lesson.getEndTime());
                    out.println(lesson.getStartWeek());
                    out.println(lesson.getEndWeek());
                    out.println(lesson.getType());
                    out.println(lesson.getCabinet());
                    out.println();
                }
                out.println("endDay");
            }
        out.flush();
    }
}
