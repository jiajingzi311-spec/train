package com.milo.sleeptimer;

import android.app.Activity;
import android.content.Intent;
import android.content.SharedPreferences;
import android.content.pm.PackageManager;
import android.content.res.ColorStateList;
import android.graphics.Color;
import android.graphics.Typeface;
import android.graphics.drawable.Drawable;
import android.graphics.drawable.GradientDrawable;
import android.graphics.drawable.RippleDrawable;
import android.os.Build;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.provider.Settings;
import android.view.Gravity;
import android.view.View;
import android.view.Window;
import android.widget.FrameLayout;
import android.widget.LinearLayout;
import android.widget.ScrollView;
import android.widget.TextView;
import android.widget.Toast;

import java.util.Locale;
import java.util.Random;

public class MainActivity extends Activity {

    private static final int[] PRESETS = {15, 30, 45, 60};
    private static final int MIN_MINUTES = 5;
    private static final int MAX_MINUTES = 180;

    private static final String[] GREETINGS = {
            "今天也辛苦啦，刷一会儿就早点睡哦",
            "放心刷吧，睡着了月亮帮你关",
            "月亮替你守着，安心闭上眼睛",
            "梦里见呀 ✨",
            "早睡的宝宝，明天皮肤会更好哦",
            "困了就睡，剩下的交给小月亮",
    };

    private static final int GOLD = 0xFFFFD98A;
    private static final int TEXT_SOFT = 0xFFC9C3E6;
    private static final int TEXT_DIM = 0xFF9C94C4;

    private SharedPreferences prefs;
    private int selectedMinutes;

    private RingView ring;
    private LinearLayout idleBox;
    private LinearLayout runningBox;
    private LinearLayout permissionCard;
    private TextView[] presetChips;
    private TextView helpBody;
    private TextView helpToggle;

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
        selectedMinutes = prefs.getInt("last_minutes", 30);

        Window win = getWindow();
        win.setStatusBarColor(Color.TRANSPARENT);
        win.setNavigationBarColor(0xFF6E4D96);
        win.getDecorView().setSystemUiVisibility(
                View.SYSTEM_UI_FLAG_LAYOUT_STABLE | View.SYSTEM_UI_FLAG_LAYOUT_FULLSCREEN);

        FrameLayout root = new FrameLayout(this);
        root.addView(new StarryView(this), new FrameLayout.LayoutParams(-1, -1));

        ScrollView scroll = new ScrollView(this);
        scroll.setVerticalScrollBarEnabled(false);
        root.addView(scroll, new FrameLayout.LayoutParams(-1, -1));

        LinearLayout col = new LinearLayout(this);
        col.setOrientation(LinearLayout.VERTICAL);
        col.setGravity(Gravity.CENTER_HORIZONTAL);
        col.setPadding(dp(24), dp(56), dp(24), dp(32));
        scroll.addView(col);

        // ---------- 标题 ----------
        TextView title = text("晚安小月亮", 30, Color.WHITE, true);
        title.setLetterSpacing(0.08f);
        col.addView(title);

        TextView greet = text(GREETINGS[new Random().nextInt(GREETINGS.length)], 14, TEXT_SOFT, false);
        greet.setPadding(0, dp(6), 0, dp(18));
        col.addView(greet);

        // ---------- 权限提示卡片 ----------
        permissionCard = new LinearLayout(this);
        permissionCard.setOrientation(LinearLayout.VERTICAL);
        permissionCard.setPadding(dp(18), dp(14), dp(18), dp(14));
        permissionCard.setBackground(round(0x33FF8FB1, 18, 0x66FF8FB1));
        TextView pTitle = text("还差一步：开启小月亮的权限", 15, Color.WHITE, true);
        pTitle.setGravity(Gravity.START);
        permissionCard.addView(pTitle);
        TextView pBody = text("点下面按钮 → 已下载的应用 → 晚安小月亮 → 打开", 13, TEXT_SOFT, false);
        pBody.setGravity(Gravity.START);
        pBody.setPadding(0, dp(4), 0, dp(10));
        permissionCard.addView(pBody);
        TextView pBtn = pill("去开启", true);
        pBtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                openAccessibilitySettings();
            }
        });
        permissionCard.addView(pBtn, new LinearLayout.LayoutParams(-1, dp(44)));
        LinearLayout.LayoutParams cardLp = new LinearLayout.LayoutParams(-1, -2);
        cardLp.bottomMargin = dp(12);
        col.addView(permissionCard, cardLp);

        // ---------- 圆环 ----------
        ring = new RingView(this);
        col.addView(ring, new LinearLayout.LayoutParams(dp(270), dp(270)));

        // ---------- 未开始时的控件 ----------
        idleBox = new LinearLayout(this);
        idleBox.setOrientation(LinearLayout.VERTICAL);
        idleBox.setGravity(Gravity.CENTER_HORIZONTAL);
        col.addView(idleBox, new LinearLayout.LayoutParams(-1, -2));

        LinearLayout stepRow = new LinearLayout(this);
        stepRow.setGravity(Gravity.CENTER);
        stepRow.setPadding(0, dp(4), 0, dp(18));
        TextView minus = circleBtn("－");
        TextView plus = circleBtn("＋");
        minus.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                setSelected(selectedMinutes - 5);
            }
        });
        plus.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                setSelected(selectedMinutes + 5);
            }
        });
        stepRow.addView(minus, new LinearLayout.LayoutParams(dp(48), dp(48)));
        TextView stepHint = text("每次 5 分钟", 13, TEXT_DIM, false);
        stepHint.setPadding(dp(20), 0, dp(20), 0);
        stepRow.addView(stepHint);
        stepRow.addView(plus, new LinearLayout.LayoutParams(dp(48), dp(48)));
        idleBox.addView(stepRow);

        LinearLayout chipRow = new LinearLayout(this);
        chipRow.setOrientation(LinearLayout.HORIZONTAL);
        presetChips = new TextView[PRESETS.length];
        for (int i = 0; i < PRESETS.length; i++) {
            final int m = PRESETS[i];
            TextView chip = pill(m + " 分钟", false);
            chip.setTextSize(14);
            chip.setOnClickListener(new View.OnClickListener() {
                @Override
                public void onClick(View v) {
                    setSelected(m);
                }
            });
            LinearLayout.LayoutParams lp = new LinearLayout.LayoutParams(0, dp(40), 1f);
            if (i > 0) {
                lp.leftMargin = dp(8);
            }
            chipRow.addView(chip, lp);
            presetChips[i] = chip;
        }
        idleBox.addView(chipRow, new LinearLayout.LayoutParams(-1, -2));

        TextView startBtn = pill("开始守护  🌙", true);
        startBtn.setTextSize(18);
        startBtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                start(selectedMinutes);
            }
        });
        LinearLayout.LayoutParams startLp = new LinearLayout.LayoutParams(-1, dp(58));
        startLp.topMargin = dp(22);
        idleBox.addView(startBtn, startLp);

        // ---------- 倒计时中的控件 ----------
        runningBox = new LinearLayout(this);
        runningBox.setOrientation(LinearLayout.VERTICAL);
        runningBox.setGravity(Gravity.CENTER_HORIZONTAL);
        col.addView(runningBox, new LinearLayout.LayoutParams(-1, -2));

        TextView runHint = text("小月亮正在守护你，放心刷吧", 14, TEXT_SOFT, false);
        runHint.setPadding(0, dp(8), 0, dp(20));
        runningBox.addView(runHint);

        TextView extendBtn = pill("再看 10 分钟", true);
        extendBtn.setTextSize(17);
        extendBtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                TimerController.extend(MainActivity.this, 10 * 60_000L);
                toast("好～再陪你 10 分钟");
                refresh();
            }
        });
        runningBox.addView(extendBtn, new LinearLayout.LayoutParams(-1, dp(54)));

        TextView cancelBtn = pill("不用了，取消守护", false);
        cancelBtn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                TimerController.cancel(MainActivity.this);
                refresh();
            }
        });
        LinearLayout.LayoutParams cancelLp = new LinearLayout.LayoutParams(-1, dp(48));
        cancelLp.topMargin = dp(12);
        runningBox.addView(cancelBtn, cancelLp);

        // ---------- 使用说明（可折叠） ----------
        LinearLayout help = new LinearLayout(this);
        help.setOrientation(LinearLayout.VERTICAL);
        help.setPadding(dp(18), dp(14), dp(18), dp(14));
        help.setBackground(round(0x1AFFFFFF, 18, 0x22FFFFFF));
        helpToggle = text("第一次用？看这里  ▾", 14, TEXT_SOFT, false);
        helpToggle.setGravity(Gravity.START);
        help.addView(helpToggle);
        helpBody = text("· 选好时间点「开始守护」，然后去刷视频就好\n"
                + "· 到点前 1 分钟会提醒，想继续看就回来点「再看 10 分钟」\n"
                + "· 到点后会自动：暂停 → 退出视频 → 回桌面 → 熄屏\n"
                + "· 守护期间下拉通知栏能看到剩余时间\n"
                + "· 自己按了锁屏键，守护会自动取消\n\n"
                + "小米手机请再设置一下，不然小月亮可能会被系统关掉：\n"
                + "1. 设置 → 应用设置 → 晚安小月亮 → 省电策略 → 无限制\n"
                + "2. 同一页 → 自启动 → 打开\n"
                + "3. 最近任务里把小月亮往下拉，锁定它\n"
                + "4. 开权限时开关是灰的：在应用信息页右上角 ⋮ → 允许受限制的设置",
                13, TEXT_SOFT, false);
        helpBody.setGravity(Gravity.START);
        helpBody.setLineSpacing(dp(3), 1f);
        helpBody.setPadding(0, dp(10), 0, 0);
        helpBody.setVisibility(View.GONE);
        help.addView(helpBody);
        help.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                boolean show = helpBody.getVisibility() != View.VISIBLE;
                helpBody.setVisibility(show ? View.VISIBLE : View.GONE);
                helpToggle.setText(show ? "第一次用？看这里  ▴" : "第一次用？看这里  ▾");
            }
        });
        LinearLayout.LayoutParams helpLp = new LinearLayout.LayoutParams(-1, -2);
        helpLp.topMargin = dp(28);
        col.addView(help, helpLp);

        TextView footer = text("不读取屏幕内容 · 不联网 · 只为让你睡个好觉 💤", 11, TEXT_DIM, false);
        footer.setPadding(0, dp(18), 0, 0);
        col.addView(footer);

        setContentView(root);
        refresh();

        // 安卓 13 以上要允许通知，才能在通知栏看到倒计时
        if (Build.VERSION.SDK_INT >= 33
                && checkSelfPermission("android.permission.POST_NOTIFICATIONS")
                != PackageManager.PERMISSION_GRANTED) {
            requestPermissions(new String[]{"android.permission.POST_NOTIFICATIONS"}, 1);
        }
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

    private void setSelected(int minutes) {
        selectedMinutes = Math.max(MIN_MINUTES, Math.min(MAX_MINUTES, minutes));
        refresh();
    }

    private void start(int minutes) {
        if (!TimerController.isAccessibilityEnabled(this)) {
            toast("先开启一下小月亮的权限哦");
            openAccessibilitySettings();
            return;
        }
        TimerController.start(this, minutes * 60_000L);
        prefs.edit().putInt("last_minutes", minutes).apply();
        toast("🌙 " + minutes + " 分钟后帮你关掉，安心刷吧～");
        refresh();
        moveTaskToBack(true);
    }

    private void refresh() {
        boolean enabled = TimerController.isAccessibilityEnabled(this);
        permissionCard.setVisibility(enabled ? View.GONE : View.VISIBLE);

        long end = TimerController.getEndTime(this);
        long left = end - System.currentTimeMillis();
        boolean running = end > 0 && left > 0;

        if (running) {
            long total = Math.max(1L, TimerController.getTotalMs(this));
            long sec = (left + 999) / 1000;
            String time = String.format(Locale.US, "%02d:%02d", sec / 60, sec % 60);
            ring.setState((float) left / total, time, "后帮你熄屏");
        } else {
            ring.setState(selectedMinutes / 60f, String.valueOf(selectedMinutes), "分钟后熄屏");
        }
        idleBox.setVisibility(running ? View.GONE : View.VISIBLE);
        runningBox.setVisibility(running ? View.VISIBLE : View.GONE);

        for (int i = 0; i < PRESETS.length; i++) {
            boolean on = PRESETS[i] == selectedMinutes;
            styleChip(presetChips[i], on);
        }
    }

    // ================= 小工具 =================

    private void openAccessibilitySettings() {
        try {
            Intent i = new Intent(Settings.ACTION_ACCESSIBILITY_SETTINGS);
            i.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
            startActivity(i);
        } catch (Exception e) {
            toast("请手动打开：设置 → 更多设置 → 无障碍");
        }
    }

    private TextView text(String s, float size, int color, boolean bold) {
        TextView t = new TextView(this);
        t.setText(s);
        t.setTextSize(size);
        t.setTextColor(color);
        t.setGravity(Gravity.CENTER);
        if (bold) {
            t.setTypeface(Typeface.DEFAULT_BOLD);
        }
        return t;
    }

    /** 圆角药丸按钮：filled=金色实心，否则半透明描边 */
    private TextView pill(String s, boolean filled) {
        TextView t = text(s, 16, filled ? 0xFF2A2457 : Color.WHITE, filled);
        t.setClickable(true);
        if (filled) {
            GradientDrawable g = new GradientDrawable(GradientDrawable.Orientation.LEFT_RIGHT,
                    new int[]{0xFFFFE29A, 0xFFFFB38A});
            g.setCornerRadius(dp(100));
            t.setBackground(ripple(g));
            t.setElevation(dp(6));
        } else {
            t.setBackground(ripple(round(0x1FFFFFFF, 100, 0x40FFFFFF)));
        }
        return t;
    }

    private void styleChip(TextView chip, boolean on) {
        if (on) {
            chip.setBackground(ripple(round(GOLD, 100, GOLD)));
            chip.setTextColor(0xFF2A2457);
            chip.setTypeface(Typeface.DEFAULT_BOLD);
        } else {
            chip.setBackground(ripple(round(0x1FFFFFFF, 100, 0x40FFFFFF)));
            chip.setTextColor(Color.WHITE);
            chip.setTypeface(Typeface.DEFAULT);
        }
    }

    private TextView circleBtn(String s) {
        TextView t = text(s, 20, Color.WHITE, false);
        t.setClickable(true);
        GradientDrawable g = new GradientDrawable();
        g.setShape(GradientDrawable.OVAL);
        g.setColor(0x1FFFFFFF);
        g.setStroke(dp(1), 0x40FFFFFF);
        t.setBackground(ripple(g));
        return t;
    }

    private GradientDrawable round(int fill, int radiusDp, int stroke) {
        GradientDrawable g = new GradientDrawable();
        g.setColor(fill);
        g.setCornerRadius(dp(radiusDp));
        g.setStroke(dp(1), stroke);
        return g;
    }

    private Drawable ripple(Drawable content) {
        return new RippleDrawable(ColorStateList.valueOf(0x40FFFFFF), content, null);
    }

    private void toast(String s) {
        Toast.makeText(this, s, Toast.LENGTH_SHORT).show();
    }

    private int dp(float v) {
        return Math.round(v * getResources().getDisplayMetrics().density);
    }
}
