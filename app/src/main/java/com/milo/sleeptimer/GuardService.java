package com.milo.sleeptimer;

import android.app.Notification;
import android.app.NotificationChannel;
import android.app.NotificationManager;
import android.app.PendingIntent;
import android.app.Service;
import android.content.Context;
import android.content.Intent;
import android.content.pm.ServiceInfo;
import android.os.Build;
import android.os.IBinder;

/**
 * 倒计时期间的常驻通知：“晚安小月亮守护中 · 还剩 xx:xx”。
 * 前台服务能防止小米把 App 冻结，同时她下拉通知栏就能看到剩余时间。
 */
public class GuardService extends Service {

    private static final String CHANNEL = "guard";
    private static final int NOTIFY_ID = 1;
    private static final String EXTRA_END = "end";
    private static final String ACTION_STOP = "com.milo.sleeptimer.STOP_GUARD";

    public static void start(Context c, long endTime) {
        Intent i = new Intent(c, GuardService.class);
        i.putExtra(EXTRA_END, endTime);
        try {
            c.startForegroundService(i);
        } catch (Exception ignored) {
            // 极少数情况下系统不让启动，闹钟依然会按时触发
        }
    }

    public static void stop(Context c) {
        try {
            c.stopService(new Intent(c, GuardService.class));
        } catch (Exception ignored) {
        }
    }

    @Override
    public int onStartCommand(Intent intent, int flags, int startId) {
        long end = intent != null ? intent.getLongExtra(EXTRA_END, 0L) : 0L;
        if (end <= 0L) {
            end = TimerController.getEndTime(this);
        }
        Notification n = buildNotification(end);
        try {
            if (Build.VERSION.SDK_INT >= 34) {
                startForeground(NOTIFY_ID, n, ServiceInfo.FOREGROUND_SERVICE_TYPE_SPECIAL_USE);
            } else {
                startForeground(NOTIFY_ID, n);
            }
        } catch (Exception e) {
            stopSelf();
            return START_NOT_STICKY;
        }
        if (end <= System.currentTimeMillis()) {
            stopSelf();
        }
        return START_NOT_STICKY;
    }

    private Notification buildNotification(long endTime) {
        NotificationManager nm = (NotificationManager) getSystemService(Context.NOTIFICATION_SERVICE);
        if (nm != null && nm.getNotificationChannel(CHANNEL) == null) {
            NotificationChannel ch = new NotificationChannel(CHANNEL, "守护中提示",
                    NotificationManager.IMPORTANCE_LOW);
            ch.setDescription("倒计时期间显示剩余时间");
            ch.setShowBadge(false);
            nm.createNotificationChannel(ch);
        }
        Intent open = new Intent(this, MainActivity.class);
        open.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK | Intent.FLAG_ACTIVITY_SINGLE_TOP);
        PendingIntent pi = PendingIntent.getActivity(this, 0, open,
                PendingIntent.FLAG_UPDATE_CURRENT | PendingIntent.FLAG_IMMUTABLE);

        Notification.Builder b = new Notification.Builder(this, CHANNEL)
                .setSmallIcon(R.drawable.ic_notify)
                .setContentTitle("🌙 晚安小月亮守护中")
                .setContentText("到点会帮你关掉视频并熄屏")
                .setContentIntent(pi)
                .setOngoing(true)
                .setOnlyAlertOnce(true)
                .setCategory(Notification.CATEGORY_ALARM);
        if (endTime > 0) {
            b.setWhen(endTime).setShowWhen(true).setUsesChronometer(true).setChronometerCountDown(true);
        }
        return b.build();
    }

    @Override
    public void onDestroy() {
        stopForeground(STOP_FOREGROUND_REMOVE);
        super.onDestroy();
    }

    @Override
    public IBinder onBind(Intent intent) {
        return null;
    }
}
