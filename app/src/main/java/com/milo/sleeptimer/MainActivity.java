package com.milo.sleeptimer;

import android.app.Activity;
import android.content.Intent;
import android.content.SharedPreferences;
import android.graphics.Color;
import android.graphics.Typeface;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.provider.Settings;
import android.text.InputType;
import android.view.Gravity;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.ScrollView;
import android.widget.TextView;
import android.widget.Toast;

import java.util.Locale;

public class MainActivity extends Activity {

    private static final int[] PRESETS = {15, 30, 45, 60};

    private TextView statusText;
    private Button permissionButton;
    private TextView countdownText;
    private Button cancelButton;
    private EditText minutesInput;
    private SharedPreferences prefs;

    private final Handler handler = new Handler(Looper.getMainLooper());
    private final Runnable ticker = new Runnable() {
        @Override
        public void run() {
            refresh();
            handler.postDelayed(this, 1000);
        }
    };

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        prefs = getSharedPreferences("settings", MODE_PRIVATE);

        int pad = dp(20);
        LinearLayout root = new LinearLayout(this);
        root.setOrientation(LinearLayout.VERTICAL);
        root.setPadding(pad, pad, pad, pad);
        root.setBackgroundColor(Color.parseColor("#FFF8F0"));

        TextView title = new TextView(this);
        title.setText("🌙 睡前定时熄屏");
        title.setTextSize(26);
        title.setTypeface(Typeface.DEFAULT_BOLD);
        title.setTextColor(Color.parseColor("#333333"));
        root.addView(title);

        TextView sub = new TextView(this);
        sub.setText("刷着视频睡着也不怕，到点自动退出并熄屏");
        sub.setTextSize(14);
        sub.setTextColor(Color.parseColor("#888888"));
        sub.setPadding(0, dp(4), 0, dp(16));
        root.addView(sub);

        statusText = new TextView(this);
        statusText.setTextSize(15);
        statusText.setPadding(0, 0, 0, dp(6));
        root.addView(statusText);

        permissionButton = new Button(this);
        permissionButton.setText("去开启权限（无障碍 → 已下载的应用 → 睡前定时熄屏）");
        permissionButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                openAccessibilitySettings();
            }
        });
        root.addView(permissionButton);

        countdownText = new TextView(this);
        countdownText.setTextSize(40);
        countdownText.setTypeface(Typeface.DEFAULT_BOLD);
        countdownText.setGravity(Gravity.CENTER);
        countdownText.setPadding(0, dp(24), 0, dp(8));
        root.addView(countdownText);

        cancelButton = new Button(this);
        cancelButton.setText("取消定时");
        cancelButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                SleepTimerService svc = SleepTimerService.getInstance();
                if (svc != null) {
                    svc.cancelTimer();
                }
                refresh();
            }
        });
        root.addView(cancelButton);

        TextView presetLabel = label("一键开始：");
        root.addView(presetLabel);

        LinearLayout presetRow = new LinearLayout(this);
        presetRow.setOrientation(LinearLayout.HORIZONTAL);
        for (final int m : PRESETS) {
            Button b = new Button(this);
            b.setText(m + "分钟");
            b.setOnClickListener(new View.OnClickListener() {
                @Override
                public void onClick(View v) {
                    start(m);
                }
            });
            presetRow.addView(b, new LinearLayout.LayoutParams(0,
                    LinearLayout.LayoutParams.WRAP_CONTENT, 1f));
        }
        root.addView(presetRow);

        root.addView(label("自定义分钟数："));

        LinearLayout customRow = new LinearLayout(this);
        customRow.setOrientation(LinearLayout.HORIZONTAL);
        minutesInput = new EditText(this);
        minutesInput.setInputType(InputType.TYPE_CLASS_NUMBER);
        minutesInput.setHint("例如 20");
        int last = prefs.getInt("last_minutes", 0);
        if (last > 0) {
            minutesInput.setText(String.valueOf(last));
        }
        customRow.addView(minutesInput, new LinearLayout.LayoutParams(0,
                LinearLayout.LayoutParams.WRAP_CONTENT, 1f));
        Button go = new Button(this);
        go.setText("开始");
        go.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                String s = minutesInput.getText().toString().trim();
                int m;
                try {
                    m = Integer.parseInt(s);
                } catch (NumberFormatException e) {
                    m = 0;
                }
                if (m <= 0 || m > 600) {
                    toast("请输入 1～600 之间的分钟数");
                    return;
                }
                start(m);
            }
        });
        customRow.addView(go);
        root.addView(customRow);

        TextView tips = new TextView(this);
        tips.setTextSize(13);
        tips.setTextColor(Color.parseColor("#888888"));
        tips.setPadding(0, dp(24), 0, 0);
        tips.setText("用法：点好时间后直接去刷视频就行。\n"
                + "· 到点前 1 分钟会弹提示，想继续看就回来重新设一次\n"
                + "· 到点后自动：暂停 → 返回 → 回桌面 → 锁屏\n"
                + "· 自己按了锁屏键，定时会自动取消\n\n"
                + "小米手机第一次用请设置：\n"
                + "1. 应用信息 → 省电策略 → 无限制\n"
                + "2. 应用信息 → 自启动 → 打开\n"
                + "3. 最近任务里把本 App 下拉“锁定”，别被清理掉");
        root.addView(tips);

        ScrollView scroll = new ScrollView(this);
        scroll.addView(root);
        setContentView(scroll);
    }

    @Override
    protected void onResume() {
        super.onResume();
        handler.post(ticker);
    }

    @Override
    protected void onPause() {
        super.onPause();
        handler.removeCallbacks(ticker);
    }

    private void start(int minutes) {
        SleepTimerService svc = SleepTimerService.getInstance();
        if (svc == null) {
            toast("请先开启无障碍权限");
            openAccessibilitySettings();
            return;
        }
        svc.startTimer(minutes * 60_000L);
        prefs.edit().putInt("last_minutes", minutes).apply();
        toast(minutes + " 分钟后自动熄屏，去刷视频吧～");
        refresh();
        moveTaskToBack(true);
    }

    private void refresh() {
        SleepTimerService svc = SleepTimerService.getInstance();
        boolean enabled = svc != null;
        if (enabled) {
            statusText.setText("✅ 权限已开启");
            statusText.setTextColor(Color.parseColor("#2E7D32"));
            permissionButton.setVisibility(View.GONE);
        } else {
            statusText.setText("⚠️ 还没开启无障碍权限，开启后才能用");
            statusText.setTextColor(Color.parseColor("#C62828"));
            permissionButton.setVisibility(View.VISIBLE);
        }

        long end = enabled ? svc.getEndTime() : 0L;
        long left = end - System.currentTimeMillis();
        if (end > 0 && left > 0) {
            long sec = left / 1000;
            countdownText.setText(String.format(Locale.US, "%02d:%02d", sec / 60, sec % 60));
            countdownText.setTextColor(Color.parseColor("#E65100"));
            cancelButton.setVisibility(View.VISIBLE);
        } else {
            countdownText.setText("未开始");
            countdownText.setTextColor(Color.parseColor("#BBBBBB"));
            cancelButton.setVisibility(View.GONE);
        }
    }

    private void openAccessibilitySettings() {
        try {
            Intent i = new Intent(Settings.ACTION_ACCESSIBILITY_SETTINGS);
            i.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
            startActivity(i);
        } catch (Exception e) {
            toast("请手动打开：设置 → 更多设置 → 无障碍");
        }
    }

    private TextView label(String text) {
        TextView t = new TextView(this);
        t.setText(text);
        t.setTextSize(15);
        t.setTextColor(Color.parseColor("#555555"));
        t.setPadding(0, dp(20), 0, dp(4));
        return t;
    }

    private void toast(String s) {
        Toast.makeText(this, s, Toast.LENGTH_SHORT).show();
    }

    private int dp(int v) {
        return Math.round(v * getResources().getDisplayMetrics().density);
    }
}
