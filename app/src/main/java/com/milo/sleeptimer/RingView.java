package com.milo.sleeptimer;

import android.content.Context;
import android.graphics.Canvas;
import android.graphics.Matrix;
import android.graphics.Paint;
import android.graphics.RectF;
import android.graphics.SweepGradient;
import android.graphics.Typeface;
import android.view.View;

/** 圆环倒计时：中间显示月亮、时间和一句小字 */
public class RingView extends View {

    private final Paint track = new Paint(Paint.ANTI_ALIAS_FLAG);
    private final Paint arc = new Paint(Paint.ANTI_ALIAS_FLAG);
    private final Paint moonP = new Paint(Paint.ANTI_ALIAS_FLAG);
    private final Paint mainP = new Paint(Paint.ANTI_ALIAS_FLAG);
    private final Paint subP = new Paint(Paint.ANTI_ALIAS_FLAG);
    private final RectF oval = new RectF();
    private final float d;

    private float progress = 1f;
    private String mainText = "30";
    private String subText = "分钟";

    public RingView(Context context) {
        super(context);
        d = getResources().getDisplayMetrics().density;
        setLayerType(LAYER_TYPE_SOFTWARE, null); // 让发光效果生效

        track.setStyle(Paint.Style.STROKE);
        track.setStrokeWidth(12 * d);
        track.setColor(0x26FFFFFF);

        arc.setStyle(Paint.Style.STROKE);
        arc.setStrokeWidth(12 * d);
        arc.setStrokeCap(Paint.Cap.ROUND);
        arc.setShadowLayer(10 * d, 0, 0, 0x99FFC58A);

        moonP.setTextAlign(Paint.Align.CENTER);
        moonP.setTextSize(30 * d);

        mainP.setTextAlign(Paint.Align.CENTER);
        mainP.setColor(0xFFFFFFFF);
        mainP.setTextSize(56 * d);
        mainP.setTypeface(Typeface.create("sans-serif-light", Typeface.NORMAL));

        subP.setTextAlign(Paint.Align.CENTER);
        subP.setColor(0xFFC9C3E6);
        subP.setTextSize(15 * d);
    }

    public void setState(float progress, String main, String sub) {
        this.progress = Math.max(0f, Math.min(1f, progress));
        this.mainText = main;
        this.subText = sub;
        invalidate();
    }

    @Override
    protected void onSizeChanged(int w, int h, int oldw, int oldh) {
        super.onSizeChanged(w, h, oldw, oldh);
        float cx = w / 2f, cy = h / 2f;
        SweepGradient g = new SweepGradient(cx, cy,
                new int[]{0xFFFFE29A, 0xFFFFB38A, 0xFFFF8FB1, 0xFFC9A0FF, 0xFFFFE29A}, null);
        Matrix m = new Matrix();
        m.setRotate(-90, cx, cy);
        g.setLocalMatrix(m);
        arc.setShader(g);
    }

    @Override
    protected void onDraw(Canvas canvas) {
        float w = getWidth(), h = getHeight();
        float pad = 12 * d + 12 * d;
        oval.set(pad, pad, w - pad, h - pad);
        canvas.drawArc(oval, 0, 360, false, track);
        if (progress > 0.001f) {
            canvas.drawArc(oval, -90, 360 * progress, false, arc);
        }
        float cx = w / 2f, cy = h / 2f;
        canvas.drawText("🌙", cx, cy - 40 * d, moonP);
        canvas.drawText(mainText, cx, cy + 22 * d, mainP);
        canvas.drawText(subText, cx, cy + 52 * d, subP);
    }
}
