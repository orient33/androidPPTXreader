package com.singuloid.officereader;

import android.annotation.SuppressLint;
import android.app.Activity;
import android.app.ProgressDialog;
import android.content.Intent;
import android.graphics.Bitmap;
import android.graphics.Bitmap.Config;
import android.graphics.Canvas;
import android.graphics.Color;
import android.graphics.Matrix;
import android.graphics.Paint;
import android.graphics.RectF;
import android.graphics.drawable.BitmapDrawable;
import android.graphics.drawable.ColorDrawable;
import android.net.Uri;
import android.os.AsyncTask;
import android.os.Bundle;
import android.os.Handler;
import android.os.Message;
import android.os.Parcelable;
import android.text.TextUtils;
import android.util.Log;
import android.view.MotionEvent;
import android.view.View;
import android.view.View.OnTouchListener;
import android.view.ViewGroup;
import android.view.ViewGroup.LayoutParams;
import android.widget.Toast;

import com.qhm123.slide.GestureDetector;
import com.qhm123.slide.ImageViewTouch;
import com.qhm123.slide.PagerAdapter;
import com.qhm123.slide.ScaleGestureDetector;
import com.qhm123.slide.ViewPager;

import net.pbdavey.awt.Graphics2D;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.atomic.AtomicBoolean;

import and.awt.Dimension;

public class OpenActivity extends Activity {

    private static final String TAG = OpenActivity.class.getSimpleName();

    private ViewPager mViewPager;
    private PagerAdapter mPagerAdapter;

    private GestureDetector mGestureDetector;
    private ScaleGestureDetector mScaleGestureDetector;

    private boolean mOnScale = false;
    private boolean mOnPagerScoll = false;

    private int slideCount = 0;
    private XSLFSlide[] slide;
    private XMLSlideShow ppt;
    private Dimension pgsize;

    private ProgressDialog mProgressDialog;

    @Override
    public void onCreate(Bundle savedInstanceState) {
        System.setProperty("javax.xml.stream.XMLInputFactory",
                "com.sun.xml.stream.ZephyrParserFactory");
        System.setProperty("javax.xml.stream.XMLOutputFactory",
                "com.sun.xml.stream.ZephyrWriterFactory");
        System.setProperty("javax.xml.stream.XMLEventFactory",
                "com.sun.xml.stream.events.ZephyrEventFactory");

        Thread.currentThread().setContextClassLoader(
                getClass().getClassLoader());

        // some test
//        try {
//            XMLInputFactory inputFactory = XMLInputFactory.newInstance();
//            XMLEventReader reader = inputFactory
//                    .createXMLEventReader(new StringReader(
//                            "<doc att=\"value\">some text</doc>"));
//            while (reader.hasNext()) {
//                XMLEvent e = reader.nextEvent();
//                Log.e("HelloStax", "Event:[" + e + "]");
//            }
//        } catch (XMLStreamException e) {
//            Log.e("HelloStax", "Error parsing XML", e);
//        }

        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_open);

        setProgressBarVisibility(true);
        setProgressBarIndeterminate(true);

        mViewPager = findViewById(R.id.pager);
        mViewPager.setPageMargin(10);
        mViewPager.setPageMarginDrawable(new ColorDrawable(Color.BLACK));
        mViewPager.setOffscreenPageLimit(1);
        mViewPager.setOnPageChangeListener(mPageChangeListener);
        setupOnTouchListeners(mViewPager);

        mProgressDialog = new ProgressDialog(this);
        mProgressDialog.setMessage("Loading");
        mProgressDialog.setIndeterminate(true);

        String path = null;
        Intent i = getIntent();
        if (i != null) {
            Uri uri = i.getData();
            if (uri != null) {
                Log.d(TAG, "uri.getPath: " + uri.getPath());
                path = uri.getPath();
            } else {
                path = i.getStringExtra("dst");//"/sdcard/talkaboutjvm.pptx";
                if (TextUtils.isEmpty(path)) {
                    path = "/sdcard/Download/example.pptx";
                }
                File demoFile = new File(path);
                if (!demoFile.exists()) {
                    Toast.makeText(this, path + " not exist!", Toast.LENGTH_LONG).show();
                    finish();
                    return;
                }
            }
        }

        try {
            setTitle(path);
            pptx2png(path);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    private void pptx2png(final String path) throws IOException,
            InvalidFormatException {

        mProgressDialog.show();

        new AsyncTask<Void, Void, Void>() {
            protected void onPostExecute(Void result) {
                mProgressDialog.dismiss();

                mViewPager.setAdapter(mPagerAdapter);
            }

            @Override
            protected Void doInBackground(Void... params) {
                System.out.println("Processing " + path);
                long time = System.currentTimeMillis();
                try {
                    ppt = new XMLSlideShow(OPCPackage.open(path,
                            PackageAccess.READ));
                    pgsize = ppt.getPageSize();
                    System.out.println("pgsize.width: " + pgsize.getWidth()
                            + ", pgsize.height: " + pgsize.getHeight());
                    slide = ppt.getSlides();
                    slideCount = slide.length;
                } catch (InvalidFormatException e) {
                    e.printStackTrace();
                }

                Log.d(TAG, "time: " + (System.currentTimeMillis() - time));
                return null;
            }
        }.execute();

        final ExecutorService es = Executors.newSingleThreadExecutor();

        final Handler handler = new Handler() {
            @Override
            public void handleMessage(Message msg) {
                super.handleMessage(msg);
                if (mViewPager == null) {
                    return;
                }
                switch (msg.what) {
                    case 0: {
                        Log.d(TAG, "draw finish");
                        View v = (View) msg.obj;
                        v.invalidate();
                        int position = msg.arg1;
                        if (position == mViewPager.getCurrentItem()) {
                            setProgress(10000);
                        }
                    }
                    break;
                    case 1: {
                        int progress = msg.arg1;
                        int max = msg.arg2;
                        int p = (int) ((float) progress / max * 10000);
                        int position = (Integer) msg.obj;
                        Log.d(TAG, "update progress: " + progress + ", max: " + max
                                + ", p: " + p + ", position: " + position);
                        if (position == 1) {
                            setProgressBarIndeterminate(false);
                        }
                        if (position == mViewPager.getCurrentItem()) {
                            if (position != 0 && progress == 0) {
                                setProgressBarIndeterminate(false);
                            }
                            setProgress(p);
                        }
                    }
                    break;
                    default:
                        break;
                }
            }
        };

        mPagerAdapter = new PagerAdapter() {

            @Override
            public void destroyItem(View container, int position, Object object) {
                ImageViewTouch view = (ImageViewTouch) object;

                view.getCanceled().set(true);
                Future<?> task = (Future<?>) view.getTag();
                task.cancel(false);

                ((ViewGroup) container).removeView(view);

                BitmapDrawable bitmapDrawable = (BitmapDrawable) view
                        .getDrawable();
                if (!bitmapDrawable.getBitmap().isRecycled()) {
                    bitmapDrawable.getBitmap().recycle();
                }

                mCache.remove(position);
            }

            @Override
            public Object instantiateItem(View container, final int position) {
                if (position == mViewPager.getCurrentItem()) {
                    setProgressBarIndeterminate(true);
                }

                final ImageViewTouch imageView = new ImageViewTouch(
                        OpenActivity.this);
                imageView.setLayoutParams(new LayoutParams(
                        LayoutParams.FILL_PARENT, LayoutParams.FILL_PARENT));
                imageView.setBackgroundColor(Color.BLACK);
                imageView.setFocusableInTouchMode(true);

                String title = slide[position].getTitle();
                System.out.println("Rendering slide " + (position + 1)
                        + (title == null ? "" : ": " + title));

                Bitmap bmp = Bitmap.createBitmap((int) pgsize.getWidth(),
                        (int) pgsize.getHeight(), Config.RGB_565);
                Canvas canvas = new Canvas(bmp);
                Paint paint = new Paint();
                paint.setColor(Color.WHITE);
                paint.setFlags(Paint.ANTI_ALIAS_FLAG);
                canvas.drawPaint(paint);

                final Graphics2D graphics2d = new Graphics2D(canvas);

                final AtomicBoolean isCanceled = new AtomicBoolean(false);
                Runnable runnable = new Runnable() {
                    @Override
                    public void run() {
                        // render
                        try {
                            slide[position].draw(graphics2d, isCanceled,
                                    handler, position);
                            handler.sendMessage(Message.obtain(handler, 0,
                                    position, 0, imageView));
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                };

                Log.d(TAG, "ViewGroup addView");

                Future<?> task = es.submit(runnable);
                imageView.setTag(task);
                imageView.setIsCanceled(isCanceled);
                imageView.setImageBitmapResetBase(bmp, true);

                ((ViewGroup) container).addView(imageView);

                Log.d(TAG, "ViewGroup addView");

                mCache.put(position, imageView);

                return imageView;
            }

            @Override
            public boolean isViewFromObject(View view, Object object) {
                return view == object;
            }

            @Override
            public int getCount() {
                return slide.length;
            }

            @Override
            public void startUpdate(View container) {
            }

            @Override
            public void finishUpdate(View container) {
            }

            @Override
            public Parcelable saveState() {
                return null;
            }

            @Override
            public void restoreState(Parcelable state, ClassLoader loader) {
            }
        };
    }

    HashMap<Integer, View> mCache = new HashMap<Integer, View>();

    public View getView(int position) {
        return mCache.get(position);
    }

    Toast mPreToast;

    ViewPager.OnPageChangeListener mPageChangeListener = new ViewPager.OnPageChangeListener() {
        @Override
        public void onPageSelected(int position, int prePosition) {
            ImageViewTouch preImageView = (ImageViewTouch) getView(prePosition);
            if (preImageView != null) {
                preImageView.setImageBitmapResetBase(
                        preImageView.mBitmapDisplayed.getBitmap(), true);
            }

            Log.d(TAG, "onPageSelected: " + position);
            if (mPreToast == null) {
                mPreToast = Toast.makeText(OpenActivity.this,
                        String.format("%d/%d", position + 1, slideCount),
                        Toast.LENGTH_SHORT);
            } else {
                mPreToast.cancel();
                mPreToast.setText(String.format("%d/%d", position + 1,
                        slideCount));
                mPreToast.setDuration(Toast.LENGTH_SHORT);
            }
            mPreToast.show();
        }

        @Override
        public void onPageScrolled(int position, float positionOffset,
                                   int positionOffsetPixels) {
            mOnPagerScoll = true;
        }

        @Override
        public void onPageScrollStateChanged(int state) {
            if (state == ViewPager.SCROLL_STATE_DRAGGING) {
                mOnPagerScoll = true;
            } else if (state == ViewPager.SCROLL_STATE_SETTLING) {
                mOnPagerScoll = false;
            } else {
                mOnPagerScoll = false;
            }
        }

    };

    public ImageViewTouch getCurrentImageView() {
        return (ImageViewTouch) getView(mViewPager.getCurrentItem());
    }

    private class MyGestureListener extends
            GestureDetector.SimpleOnGestureListener {

        @Override
        public boolean onScroll(MotionEvent e1, MotionEvent e2,
                                float distanceX, float distanceY) {
            // Logger.d(TAG, "gesture onScroll");
            if (mOnScale) {
                return true;
            }
            ImageViewTouch imageView = getCurrentImageView();
            if (imageView != null) {
                imageView.panBy(-distanceX, -distanceY);

                // 超出边界效果去掉这个
                imageView.center(true, true);
            }

            return true;
        }

        @Override
        public boolean onSingleTapConfirmed(MotionEvent e) {
            return true;
        }

        @Override
        public boolean onDoubleTap(MotionEvent e) {
            ImageViewTouch imageView = getCurrentImageView();
            // Switch between the original scale and 3x scale.
            if (imageView.mBaseZoom < 1) {
                if (imageView.getScale() > 2F) {
                    imageView.zoomTo(1f);
                } else {
                    imageView.zoomToPoint(3f, e.getX(), e.getY());
                }
            } else {
                if (imageView.getScale() > (imageView.mMinZoom + imageView.mMaxZoom) / 2f) {
                    imageView.zoomTo(imageView.mMinZoom);
                } else {
                    imageView.zoomToPoint(imageView.mMaxZoom, e.getX(),
                            e.getY());
                }
            }

            return true;
        }
    }

    private class MyOnScaleGestureListener extends
            ScaleGestureDetector.SimpleOnScaleGestureListener {

        float currentScale;
        float currentMiddleX;
        float currentMiddleY;

        @Override
        public void onScaleEnd(ScaleGestureDetector detector) {

            final ImageViewTouch imageView = getCurrentImageView();

            Log.d(TAG, "currentScale: " + currentScale + ", maxZoom: "
                    + imageView.mMaxZoom);
            if (currentScale > imageView.mMaxZoom) {
                imageView
                        .zoomToNoCenterWithAni(currentScale
                                        / imageView.mMaxZoom, 1, currentMiddleX,
                                currentMiddleY);
                currentScale = imageView.mMaxZoom;
                imageView.zoomToNoCenterValue(currentScale, currentMiddleX,
                        currentMiddleY);
            } else if (currentScale < imageView.mMinZoom) {
                // imageView.zoomToNoCenterWithAni(currentScale,
                // imageView.mMinZoom, currentMiddleX, currentMiddleY);
                currentScale = imageView.mMinZoom;
                imageView.zoomToNoCenterValue(currentScale, currentMiddleX,
                        currentMiddleY);
            } else {
                imageView.zoomToNoCenter(currentScale, currentMiddleX,
                        currentMiddleY);
            }

            imageView.center(true, true);

            // NOTE: 延迟修正缩放后可能移动问题
            imageView.postDelayed(new Runnable() {
                @Override
                public void run() {
                    mOnScale = false;
                }
            }, 300);
            // Logger.d(TAG, "gesture onScaleEnd");
        }

        @Override
        public boolean onScaleBegin(ScaleGestureDetector detector) {
            // Logger.d(TAG, "gesture onScaleStart");
            mOnScale = true;
            return true;
        }

        @Override
        public boolean onScale(ScaleGestureDetector detector, float mx, float my) {
            // Logger.d(TAG, "gesture onScale");
            ImageViewTouch imageView = getCurrentImageView();
            float ns = imageView.getScale() * detector.getScaleFactor();

            currentScale = ns;
            currentMiddleX = mx;
            currentMiddleY = my;

            if (detector.isInProgress()) {
                imageView.zoomToNoCenter(ns, mx, my);
            }
            return true;
        }
    }

    private void setupOnTouchListeners(View rootView) {
        mGestureDetector = new GestureDetector(this, new MyGestureListener(),
                null, true);
        mScaleGestureDetector = new ScaleGestureDetector(this,
                new MyOnScaleGestureListener());

        OnTouchListener rootListener = new OnTouchListener() {
            @SuppressLint("ClickableViewAccessibility")
            public boolean onTouch(View v, MotionEvent event) {
                // NOTE: gestureDetector may handle onScroll..
                if (!mOnScale) {
                    if (!mOnPagerScoll) {
                        try {
                            mGestureDetector.onTouchEvent(event);
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                }
                if (!mOnPagerScoll) {
                    mScaleGestureDetector.onTouchEvent(event);
                }

                ImageViewTouch imageView = getCurrentImageView();
                if (!mOnScale && imageView != null
                        && imageView.mBitmapDisplayed.getBitmap() != null) {
                    Matrix m = imageView.getImageViewMatrix();
                    RectF rect = new RectF(0, 0, imageView.mBitmapDisplayed
                            .getBitmap().getWidth(), imageView.mBitmapDisplayed
                            .getBitmap().getHeight());
                    m.mapRect(rect);
                    // Logger.d(TAG, "rect.right: " + rect.right +
                    // ", rect.left: "
                    // + rect.left + ", imageView.getWidth(): "
                    // + imageView.getWidth());
                    // 图片超出屏幕范围后移动
                    if (!(rect.right > imageView.getWidth() + 0.1 && rect.left < -0.1)) {
                        try {
                            mViewPager.onTouchEvent(event);
                        } catch (Exception e) {
                            // why?
                            e.printStackTrace();
                        }
                    }
                }

                // We do not use the return value of
                // mGestureDetector.onTouchEvent because we will not receive
                // the "up" event if we return false for the "down" event.
                return true;
            }
        };

        rootView.setOnTouchListener(rootListener);
    }

    @Override
    protected void onDestroy() {
        super.onDestroy();
        if (mProgressDialog != null && mProgressDialog.isShowing()) {
            mProgressDialog.dismiss();
        }
        ImageViewTouch imageView = getCurrentImageView();
        if (imageView != null) {
            imageView.mBitmapDisplayed.recycle();
            imageView.clear();
        }

        ppt = null;
        slide = null;
        mPagerAdapter = null;
        mViewPager = null;
    }
}
