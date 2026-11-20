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

/**
 * 无障碍服务：只负责“动手”——暂停、返回、回桌面、锁屏。
 * 什么时候动手由 TimerController 用系统闹钟决定。
 */
public class SleepTimerService extends AccessibilityService {

    private static SleepTimerService instance;

    public static SleepTimerService getInstance() {
        return instance;
    }

    private final Handler handler = new Handler(Looper.getMainLooper());

    // 她自己按了锁屏键 → 说明已经不看了，自动取消守护
    private final BroadcastReceiver screenOffReceiver = new BroadcastReceiver() {
        @Override
        public void onReceive(Context context, Intent intent) {
            if (Intent.ACTION_SCREEN_OFF.equals(intent.getAction()) && !shuttingDown) {
                TimerController.cancel(getApplicationContext());
            }
        }
    };
    private boolean receiverRegistered = false;
    private boolean shuttingDown = false;

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
        // 不读取任何屏幕内容
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

    /** 暂停 → 返回 → 返回 → 回桌面 → 锁屏，做完后调用 onDone */
    public void performShutdown(final Runnable onDone) {
        shuttingDown = true;
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
        handler.postDelayed(new Runnable() {
            @Override
            public void run() {
                shuttingDown = false;
                if (onDone != null) {
                    onDone.run();
                }
            }
        }, 2800);
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
