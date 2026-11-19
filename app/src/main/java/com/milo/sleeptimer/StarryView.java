package com.milo.sleeptimer;

import android.content.Context;
import android.graphics.Canvas;
import android.graphics.LinearGradient;
import android.graphics.Paint;
import android.graphics.Shader;
import android.view.View;

import java.util.Random;

/** 夜空背景：渐变 + 会轻轻闪烁的星星 */
public class StarryView extends View {

    private static final int STAR_COUNT = 70;

    private final Paint bgPaint = new Paint();
    private final Paint starPaint = new Paint(Paint.ANTI_ALIAS_FLAG);
    private float[] sx, sy, sr, sp, ss;
    private boolean running = false;

    private final Runnable tick = new Runnable() {
        @Override
        public void run() {
            invalidate();
            if (running) {
                postDelayed(this, 80);
            }
        }
    };

    public StarryView(Context context) {
        super(context);
        starPaint.setColor(0xFFFFFFFF);
    }

    @Override
    protected void onSizeChanged(int w, int h, int oldw, int oldh) {
        super.onSizeChanged(w, h, oldw, oldh);
        bgPaint.setShader(new LinearGradient(0, 0, 0, h,
                new int[]{0xFF141233, 0xFF2A2457, 0xFF4E3A7A, 0xFF7A5596},
                new float[]{0f, 0.35f, 0.75f, 1f}, Shader.TileMode.CLAMP));
        float d = getResources().getDisplayMetrics().density;
        Random rnd = new Random(520);
        sx = new float[STAR_COUNT];
        sy = new float[STAR_COUNT];
        sr = new float[STAR_COUNT];
        sp = new float[STAR_COUNT];
        ss = new float[STAR_COUNT];
        for (int i = 0; i < STAR_COUNT; i++) {
            sx[i] = rnd.nextFloat() * w;
            // 星星主要集中在上半部分
            sy[i] = (float) Math.pow(rnd.nextFloat(), 1.6) * h * 0.85f;
            sr[i] = (0.6f + rnd.nextFloat() * 1.4f) * d;
            sp[i] = rnd.nextFloat() * 6.28f;
            ss[i] = 0.6f + rnd.nextFloat() * 1.6f;
        }
    }

    @Override
    protected void onDraw(Canvas canvas) {
        canvas.drawRect(0, 0, getWidth(), getHeight(), bgPaint);
        if (sx == null) {
            return;
        }
        double t = System.currentTimeMillis() / 1000.0;
        for (int i = 0; i < STAR_COUNT; i++) {
            double a = 0.55 + 0.45 * Math.sin(t * ss[i] + sp[i]);
            starPaint.setAlpha((int) (40 + 200 * a));
            canvas.drawCircle(sx[i], sy[i], sr[i], starPaint);
        }
    }

    @Override
    protected void onAttachedToWindow() {
        super.onAttachedToWindow();
        running = true;
        post(tick);
    }

    @Override
    protected void onDetachedFromWindow() {
        running = false;
        removeCallbacks(tick);
        super.onDetachedFromWindow();
    }
}
