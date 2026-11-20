package com.milo.sleeptimer;

import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.widget.Toast;

/** 系统闹钟到点后会叫醒这里，即使 App 在后台被冻结了也一样 */
public class AlarmReceiver extends BroadcastReceiver {

    @Override
    public void onReceive(final Context context, Intent intent) {
        String action = intent.getAction();
        final Context app = context.getApplicationContext();

        if (TimerController.ACTION_WARN.equals(action)) {
            if (TimerController.isRunning(app)) {
                Toast.makeText(app, "🌙 还有 1 分钟就帮你熄屏啦，困了就睡吧～", Toast.LENGTH_LONG).show();
            }
            return;
        }

        if (TimerController.ACTION_FIRE.equals(action)) {
            long end = TimerController.getEndTime(app); // 超时太久会返回 0
            if (end == 0L) {
                GuardService.stop(app);
                return;
            }
            TimerController.clearState(app);
            final PendingResult result = goAsync(); // 让 App 多醒几秒，把关闭动作做完
            SleepTimerService svc = SleepTimerService.getInstance();
            if (svc != null) {
                svc.performShutdown(new Runnable() {
                    @Override
                    public void run() {
                        GuardService.stop(app);
                        result.finish();
                    }
                });
            } else {
                Toast.makeText(app, "小月亮的无障碍权限被系统关掉了，请重新开启", Toast.LENGTH_LONG).show();
                GuardService.stop(app);
                result.finish();
            }
        }
    }
}
