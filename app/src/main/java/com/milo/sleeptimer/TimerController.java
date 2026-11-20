package com.milo.sleeptimer;

import android.app.AlarmManager;
import android.app.PendingIntent;
import android.content.ComponentName;
import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.os.Build;
import android.provider.Settings;

/**
 * 定时的“大脑”。
 * 时间保存在本地，用系统闹钟（AlarmManager）到点叫醒 App，
 * 这样即使小米把后台 App 冻结了，到点也一定会执行。
 */
public final class TimerController {

    static final String ACTION_WARN = "com.milo.sleeptimer.WARN";
    static final String ACTION_FIRE = "com.milo.sleeptimer.FIRE";

    private static final long WARN_BEFORE_MS = 60_000L;
    /** 超时太久（比如手机睡过了）就不执行，免得第二天一拿起手机就被锁 */
    static final long MAX_LATE_MS = 5 * 60_000L;

    private static final String PREFS = "timer_state";
    private static final String KEY_END = "end";
    private static final String KEY_TOTAL = "total";

    private TimerController() {
    }

    private static SharedPreferences prefs(Context c) {
        return c.getApplicationContext().getSharedPreferences(PREFS, Context.MODE_PRIVATE);
    }

    /** 0 表示没有在倒计时 */
    public static long getEndTime(Context c) {
        long end = prefs(c).getLong(KEY_END, 0L);
        if (end > 0 && System.currentTimeMillis() - end > MAX_LATE_MS) {
            clearState(c);
            return 0L;
        }
        return end;
    }

    public static long getTotalMs(Context c) {
        return prefs(c).getLong(KEY_TOTAL, 0L);
    }

    public static boolean isRunning(Context c) {
        return getEndTime(c) > System.currentTimeMillis();
    }

    public static void start(Context c, long durationMs) {
        schedule(c, System.currentTimeMillis() + durationMs, durationMs);
    }

    public static void extend(Context c, long extraMs) {
        long end = getEndTime(c);
        long now = System.currentTimeMillis();
        if (end <= now) {
            return;
        }
        schedule(c, end + extraMs, getTotalMs(c) + extraMs);
    }

    public static void cancel(Context c) {
        cancelAlarms(c);
        clearState(c);
        GuardService.stop(c);
    }

    static void clearState(Context c) {
        prefs(c).edit().remove(KEY_END).remove(KEY_TOTAL).apply();
    }

    private static void schedule(Context c, long endTime, long totalMs) {
        cancelAlarms(c);
        prefs(c).edit().putLong(KEY_END, endTime).putLong(KEY_TOTAL, totalMs).apply();
        long now = System.currentTimeMillis();
        if (endTime - now > WARN_BEFORE_MS + 5_000L) {
            setAlarm(c, endTime - WARN_BEFORE_MS, pending(c, ACTION_WARN, 1));
        }
        setAlarm(c, endTime, pending(c, ACTION_FIRE, 2));
        GuardService.start(c, endTime);
    }

    private static void setAlarm(Context c, long at, PendingIntent pi) {
        AlarmManager am = (AlarmManager) c.getSystemService(Context.ALARM_SERVICE);
        if (am == null) {
            return;
        }
        try {
            if (Build.VERSION.SDK_INT >= 31 && !am.canScheduleExactAlarms()) {
                am.setAndAllowWhileIdle(AlarmManager.RTC_WAKEUP, at, pi);
            } else {
                am.setExactAndAllowWhileIdle(AlarmManager.RTC_WAKEUP, at, pi);
            }
        } catch (SecurityException e) {
            am.setAndAllowWhileIdle(AlarmManager.RTC_WAKEUP, at, pi);
        }
    }

    private static void cancelAlarms(Context c) {
        AlarmManager am = (AlarmManager) c.getSystemService(Context.ALARM_SERVICE);
        if (am == null) {
            return;
        }
        am.cancel(pending(c, ACTION_WARN, 1));
        am.cancel(pending(c, ACTION_FIRE, 2));
    }

    private static PendingIntent pending(Context c, String action, int code) {
        Intent i = new Intent(c, AlarmReceiver.class);
        i.setAction(action);
        return PendingIntent.getBroadcast(c, code, i,
                PendingIntent.FLAG_UPDATE_CURRENT | PendingIntent.FLAG_IMMUTABLE);
    }

    /** 直接读系统设置判断无障碍权限有没有开，比看服务对象更可靠 */
    public static boolean isAccessibilityEnabled(Context c) {
        String enabled = Settings.Secure.getString(c.getContentResolver(),
                Settings.Secure.ENABLED_ACCESSIBILITY_SERVICES);
        if (enabled == null) {
            return false;
        }
        String me = new ComponentName(c, SleepTimerService.class).flattenToString();
        String meShort = new ComponentName(c, SleepTimerService.class).flattenToShortString();
        for (String s : enabled.split(":")) {
            if (s.equalsIgnoreCase(me) || s.equalsIgnoreCase(meShort)) {
                return true;
            }
        }
        return false;
    }
}
