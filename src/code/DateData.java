package code;

import java.time.Duration;
import java.time.LocalDateTime;
import java.time.Month;

public class DateData {

    private static int WEEK_NUMBER;
    private static LocalDateTime now;


    public static void setDate(){
        now = LocalDateTime.now();
        LocalDateTime then = LocalDateTime.of(2015, Month.SEPTEMBER, 1, 0, 0, 0);
        Duration duration = Duration.between(then, now);
        WEEK_NUMBER = (int) Math.abs((duration.toDays()+1)/7.0)+1;
    }



    public static int getDayOfWeek(){
        return now.getDayOfWeek().getValue();
    }

    public static void printDate(){
        System.out.println(now.toString());
    }

    public static int getDayOfMonth(){
        return now.getDayOfMonth();
    }

    public static int getWeekNumber() {
        return WEEK_NUMBER;
    }
}