package com.milo.sleeptimer;

import android.accessibilityservice.AccessibilityService;
import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.content.IntentFilter;
import android.media.AudioAttributes;
import android.media.AudioFocusRequest;
import android.media.AudioManager;
import android.os.Handler;
import android.os.Looper;
import android.view.KeyEvent;
import android.view.accessibility.AccessibilityEvent;
import android.view.accessibility.AccessibilityNodeInfo;
import android.view.accessibility.AccessibilityWindowInfo;

import java.util.ArrayDeque;
import java.util.List;

/**
 * 无障碍服务：只负责“动手”——暂停、返回、回桌面、关掉视频小窗、锁屏。
 * 什么时候动手由 TimerController 用系统闹钟决定。
 */
public class SleepTimerService extends AccessibilityService {

    private static SleepTimerService instance;

    public static SleepTimerService getInstance() {
        return instance;
    }

    private final Handler handler = new Handler(Looper.getMainLooper());

    private static AudioFocusRequest focusRequest;

    // 她解锁手机（说明醒了）→ 把熄屏时调成 0 的音量恢复回来
    private final BroadcastReceiver unlockReceiver = new BroadcastReceiver() {
        @Override
        public void onReceive(Context context, Intent intent) {
            if (Intent.ACTION_USER_PRESENT.equals(intent.getAction())) {
                restoreVolume(getApplicationContext());
            }
        }
    };

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
            registerReceiver(unlockReceiver, new IntentFilter(Intent.ACTION_USER_PRESENT));
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
                unregisterReceiver(unlockReceiver);
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
        grabAudioFocus();   // 抢走播放权：微信直播/视频会像来电话一样停下
        pauseMedia();       // 普通视频再补一个“暂停”
        muteMedia();        // 最后保险：媒体音量调 0，直播还在小窗里放也听不到
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
        // 回到桌面后，微信直播常会变成小窗继续播 → 找到小窗的“关闭”按钮点掉（试两次，等动画结束）
        handler.postDelayed(new Runnable() {
            @Override
            public void run() {
                closeFloatingWindows();
            }
        }, 1900);
        handler.postDelayed(new Runnable() {
            @Override
            public void run() {
                closeFloatingWindows();
            }
        }, 2700);
        handler.postDelayed(new Runnable() {
            @Override
            public void run() {
                performGlobalAction(GLOBAL_ACTION_LOCK_SCREEN);
            }
        }, 3300);
        handler.postDelayed(new Runnable() {
            @Override
            public void run() {
                shuttingDown = false;
                if (onDone != null) {
                    onDone.run();
                }
            }
        }, 4000);
    }

    /** 只看会留小窗的视频 App（微信等）的窗口，找到“关闭”按钮就点 */
    private static final String[] FLOAT_APPS = {
            "com.tencent.mm",            // 微信（视频号 / 直播）
            "com.ss.android.ugc.aweme",  // 抖音
            "com.smile.gifmaker",        // 快手
    };

    private void closeFloatingWindows() {
        List<AccessibilityWindowInfo> windows;
        try {
            windows = getWindows();
        } catch (Exception e) {
            return;
        }
        if (windows == null) {
            return;
        }
        for (AccessibilityWindowInfo w : windows) {
            AccessibilityNodeInfo root;
            try {
                root = w.getRoot();
            } catch (Exception e) {
                continue;
            }
            if (root == null || root.getPackageName() == null) {
                continue;
            }
            if (!isFloatApp(root.getPackageName().toString())) {
                continue;
            }
            AccessibilityNodeInfo close = findCloseButton(root);
            if (close != null) {
                clickNodeOrParent(close);
            }
        }
    }

    private static boolean isFloatApp(String pkg) {
        for (String p : FLOAT_APPS) {
            if (p.equals(pkg)) {
                return true;
            }
        }
        return false;
    }

    private static AccessibilityNodeInfo findCloseButton(AccessibilityNodeInfo root) {
        ArrayDeque<AccessibilityNodeInfo> queue = new ArrayDeque<>();
        queue.add(root);
        int guard = 0;
        while (!queue.isEmpty() && guard++ < 400) {
            AccessibilityNodeInfo n = queue.poll();
            if (n == null) {
                continue;
            }
            if (looksLikeClose(n)) {
                return n;
            }
            for (int i = 0; i < n.getChildCount(); i++) {
                AccessibilityNodeInfo c = n.getChild(i);
                if (c != null) {
                    queue.add(c);
                }
            }
        }
        return null;
    }

    private static boolean looksLikeClose(AccessibilityNodeInfo n) {
        String desc = n.getContentDescription() == null ? "" : n.getContentDescription().toString();
        String text = n.getText() == null ? "" : n.getText().toString();
        String id = n.getViewIdResourceName() == null ? "" : n.getViewIdResourceName().toLowerCase();
        String label = (desc + "|" + text).trim();
        if (label.contains("关闭") || label.contains("退出直播") || label.contains("结束播放")) {
            return true;
        }
        if (text.equals("×") || text.equals("✕") || desc.equalsIgnoreCase("close")) {
            return true;
        }
        return id.endsWith("/close") || id.contains("close_btn") || id.contains("btn_close")
                || id.contains("float_close") || id.contains("close_icon") || id.contains("iv_close");
    }

    private static void clickNodeOrParent(AccessibilityNodeInfo n) {
        AccessibilityNodeInfo cur = n;
        for (int i = 0; i < 4 && cur != null; i++) {
            if (cur.isClickable()) {
                cur.performAction(AccessibilityNodeInfo.ACTION_CLICK);
                return;
            }
            cur = cur.getParent();
        }
        n.performAction(AccessibilityNodeInfo.ACTION_CLICK);
    }

    private void grabAudioFocus() {
        try {
            AudioManager am = (AudioManager) getSystemService(Context.AUDIO_SERVICE);
            if (am == null) {
                return;
            }
            if (focusRequest != null) {
                am.abandonAudioFocusRequest(focusRequest);
            }
            focusRequest = new AudioFocusRequest.Builder(AudioManager.AUDIOFOCUS_GAIN)
                    .setAudioAttributes(new AudioAttributes.Builder()
                            .setUsage(AudioAttributes.USAGE_MEDIA)
                            .setContentType(AudioAttributes.CONTENT_TYPE_MOVIE)
                            .build())
                    .setOnAudioFocusChangeListener(new AudioManager.OnAudioFocusChangeListener() {
                        @Override
                        public void onAudioFocusChange(int focusChange) {
                            // 别的 App 再要回去就让给它，什么都不用做
                        }
                    })
                    .build();
            am.requestAudioFocus(focusRequest);
        } catch (Exception ignored) {
        }
    }

    private void muteMedia() {
        try {
            AudioManager am = (AudioManager) getSystemService(Context.AUDIO_SERVICE);
            if (am == null) {
                return;
            }
            int vol = am.getStreamVolume(AudioManager.STREAM_MUSIC);
            if (vol > 0) {
                getApplicationContext().getSharedPreferences("timer_state", Context.MODE_PRIVATE)
                        .edit().putInt("saved_volume", vol).apply();
                am.setStreamVolume(AudioManager.STREAM_MUSIC, 0, 0);
            }
        } catch (Exception ignored) {
        }
    }

    /** 把熄屏时静音前的音量恢复（只在音量还是 0 的时候恢复，不覆盖她自己调过的） */
    public static void restoreVolume(Context c) {
        try {
            android.content.SharedPreferences p =
                    c.getSharedPreferences("timer_state", Context.MODE_PRIVATE);
            int saved = p.getInt("saved_volume", -1);
            if (saved <= 0) {
                return;
            }
            p.edit().remove("saved_volume").apply();
            AudioManager am = (AudioManager) c.getSystemService(Context.AUDIO_SERVICE);
            if (am != null && am.getStreamVolume(AudioManager.STREAM_MUSIC) == 0) {
                am.setStreamVolume(AudioManager.STREAM_MUSIC, saved, 0);
            }
        } catch (Exception ignored) {
        }
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
