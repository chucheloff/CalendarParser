package code;


import biweekly.ICalVersion;
import biweekly.ICalendar;
import biweekly.component.VAlarm;
import biweekly.component.VEvent;
import biweekly.io.text.ICalWriter;
import biweekly.parameter.Related;
import biweekly.property.Trigger;
import biweekly.util.Duration;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.Month;
import java.util.HashMap;
import java.util.TimeZone;


public class CalendarCreator {


    private static ICalendar ical;
    private static VEvent event;
    private static DateFormat df;
    private static HashMap<Integer, Integer> monthsLenght;


    /**
     * Класс, создающий .ics файлы на основе данных массива lessonsData
     */
    public static void create() throws IOException , ParseException {
        ical = new ICalendar();

        df = new SimpleDateFormat("yyyy-MM-dd HH:mm");
        ical.setProductId("Chuchelov's Calendar");

        LocalDate then = LocalDate.of(2015, Month.AUGUST, 31);
        LocalDate date;
        for (int dayData = 0; dayData < 6; dayData++) {
            for (Lesson lesson : TablesReader.getDayData(dayData).getDATA()) {
                date = then;
                // generating lessons for weeks it appears
                //System.out.println(lesson.getLessonName() +" : ");
                System.out.println("\nweeks: "+lesson.getStartWeek() +" - "+ lesson.getEndWeek());
                date = date.plusWeeks(lesson.getStartWeek()-2);
                for (int i = lesson.getStartWeek()-1; i <lesson.getEndWeek(); i++) {
                    date = date.plusWeeks(1);
                    System.out.println(lesson.getLessonName() + " on " + date.toString() + " at " +
                            lesson.getStartTime() + " - " + lesson.getEndTime() + " by "
                            + lesson.getLecturer());

                    //Creating event
                    event = new VEvent();
                    event.setSummary(lesson.getLessonName());
                    event.setLocation(lesson.getCabinet());
                    event.setDateStart(df.parse(date + " " + lesson.getStartTime().replace("-", ":").replace(".",":")));
                    event.setDateEnd(df.parse(date + " " + lesson.getEndTime().replace("-", ":").replace(".",":")));
                    if (lesson.getLecturer() != null)
                        event.setDescription(
                                lesson.getType()+ "\n" + lesson.getLecturer());
                    else event.setDescription(lesson.getType());
                    Duration duration = Duration.builder().prior(true).minutes(10).build();
                    Trigger trigger = new Trigger(duration, Related.START);
                    VAlarm alarm = VAlarm.display(trigger, lesson.getLessonName() + " через 10 мин");
                    event.addAlarm(alarm);
                    ical.addEvent(event);
                }
            }
            then = then.plusDays(1);
        }
        File file = new File("src/res/cals/Schedule_"+TablesReader.group+".ics");

        ICalWriter writer = new ICalWriter(new FileOutputStream(file), ICalVersion.V2_0);

        TimeZone tz = TimeZone.getTimeZone("Europe/Moscow");
        writer.getTimezoneInfo().setDefaultTimeZone(tz);

        writer.write(ical);
        writer.flush();
        writer.close();

        System.out.println("Writing succeed");

    }
}