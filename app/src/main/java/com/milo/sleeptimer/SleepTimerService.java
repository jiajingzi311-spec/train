package com.milo.sleeptimer;

import android.accessibilityservice.AccessibilityService;
import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.content.IntentFilter;
import android.media.AudioManager;
import android.os.Handler;
import android.os.Looper;
import android.view.KeyEvent;
import android.view.accessibility.AccessibilityEvent;
import android.widget.Toast;

/**
 * 核心服务：倒计时结束后 → 暂停媒体 → 返回 → 回到桌面 → 锁屏熄屏。
 * 用无障碍服务实现，因为只有它能在别的 App（微信视频号）上面执行“返回 / 桌面 / 锁屏”。
 */
public class SleepTimerService extends AccessibilityService {

    private static final long WARN_BEFORE_MS = 60_000L;      // 提前 1 分钟提醒
    private static final long MAX_LATE_MS = 5 * 60_000L;      // 超时太久（手机睡过了）就不执行

    private static SleepTimerService instance;

    public static SleepTimerService getInstance() {
        return instance;
    }

    private final Handler handler = new Handler(Looper.getMainLooper());
    private long endTime = 0L;
    private long totalMs = 0L;

    private final Runnable fireRunnable = new Runnable() {
        @Override
        public void run() {
            fire();
        }
    };

    private final Runnable warnRunnable = new Runnable() {
        @Override
        public void run() {
            Toast.makeText(SleepTimerService.this,
                    "🌙 还有 1 分钟就帮你熄屏啦，困了就睡吧～", Toast.LENGTH_LONG).show();
        }
    };

    // 她自己按了锁屏键 → 说明已经不看了，自动取消定时，避免第二天一拿起手机就被锁
    private final BroadcastReceiver screenOffReceiver = new BroadcastReceiver() {
        @Override
        public void onReceive(Context context, Intent intent) {
            if (Intent.ACTION_SCREEN_OFF.equals(intent.getAction())) {
                cancelTimer();
            }
        }
    };
    private boolean receiverRegistered = false;

    @Override
    protected void onServiceConnected() {
        super.onServiceConnected();
        instance = this;
        if (!receiverRegistered) {
            registerReceiver(screenOffReceiver, new IntentFilter(Intent.ACTION_SCREEN_OFF));
            receiverRegistered = true;
        }
    }

    @Override
    public void onAccessibilityEvent(AccessibilityEvent event) {
        // 不需要读取任何屏幕内容
    }

    @Override
    public void onInterrupt() {
    }

    @Override
    public boolean onUnbind(Intent intent) {
        cleanup();
        return super.onUnbind(intent);
    }

    @Override
    public void onDestroy() {
        cleanup();
        super.onDestroy();
    }

    private void cleanup() {
        cancelTimer();
        if (receiverRegistered) {
            try {
                unregisterReceiver(screenOffReceiver);
            } catch (IllegalArgumentException ignored) {
            }
            receiverRegistered = false;
        }
        if (instance == this) {
            instance = null;
        }
    }

    public void startTimer(long durationMs) {
        cancelTimer();
        endTime = System.currentTimeMillis() + durationMs;
        totalMs = durationMs;
        handler.postDelayed(fireRunnable, durationMs);
        if (durationMs > WARN_BEFORE_MS) {
            handler.postDelayed(warnRunnable, durationMs - WARN_BEFORE_MS);
        }
    }

    public void cancelTimer() {
        handler.removeCallbacks(fireRunnable);
        handler.removeCallbacks(warnRunnable);
        endTime = 0L;
    }

    /** 在当前倒计时基础上再加一段时间 */
    public void extendTimer(long extraMs) {
        if (endTime == 0L) {
            return;
        }
        long left = Math.max(0L, endTime - System.currentTimeMillis());
        long newTotal = totalMs + extraMs;
        startTimer(left + extraMs);
        totalMs = newTotal;
    }

    public long getTotalMs() {
        return totalMs;
    }

    /** 0 表示没有在倒计时 */
    public long getEndTime() {
        return endTime;
    }

    private void fire() {
        long late = System.currentTimeMillis() - endTime;
        endTime = 0L;
        if (late > MAX_LATE_MS) {
            return;
        }
        pauseMedia();
        performGlobalAction(GLOBAL_ACTION_BACK);
        handler.postDelayed(new Runnable() {
            @Override
            public void run() {
                performGlobalAction(GLOBAL_ACTION_BACK);
            }
        }, 500);
        handler.postDelayed(new Runnable() {
            @Override
            public void run() {
                performGlobalAction(GLOBAL_ACTION_HOME);
            }
        }, 1100);
        handler.postDelayed(new Runnable() {
            @Override
            public void run() {
                performGlobalAction(GLOBAL_ACTION_LOCK_SCREEN);
            }
        }, 2000);
    }

    private void pauseMedia() {
        try {
            AudioManager am = (AudioManager) getSystemService(Context.AUDIO_SERVICE);
            if (am != null) {
                am.dispatchMediaKeyEvent(new KeyEvent(KeyEvent.ACTION_DOWN, KeyEvent.KEYCODE_MEDIA_PAUSE));
                am.dispatchMediaKeyEvent(new KeyEvent(KeyEvent.ACTION_UP, KeyEvent.KEYCODE_MEDIA_PAUSE));
            }
        } catch (Exception ignored) {
        }
    }
}
