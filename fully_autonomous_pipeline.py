import os
import re
import subprocess
import textwrap
import time
from pathlib import Path

import numpy as np
import pandas as pd
import tifffile
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt


# =============================================================================
# 1. CONFIGURATION
# =============================================================================

IMAGE_FOLDER = r"C:\Users\udeer\Downloads\Filopodia analysis\Filopodia analysis"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "filopodyan_output")
TABLE_DIR = os.path.join(OUTPUT_DIR, "Filopodyan_tables")
THUMBNAIL_DIR = os.path.join(OUTPUT_DIR, "thumbnails")
LOG_DIR = os.path.join(OUTPUT_DIR, "logs")

OUTPUT_PPTX = os.path.join(
    SCRIPT_DIR,
    "detailed_filopodyan_comprehensive_report.pptx",
)

OUTPUT_SUMMARY_CSV = os.path.join(
    OUTPUT_DIR,
    "filopodyan_run_summary.csv",
)

# EXACT installed Fiji executable. Nothing is downloaded or installed.
FIJI_EXE = r"C:\Users\udeer\Downloads\fiji-latest-win64-jdk\Fiji\fiji-windows-x64.exe"

# The uploaded/installed Filopodyan JAR is already inside this Fiji.
# This path is informational only; the Python script does NOT copy or install it.
FILOPODYAN_JAR_NAME = "Filopodyan_.jar"

# Give one movie this long before declaring the Fiji instance failed.
FILOPODYAN_TIMEOUT_SECONDS = 30 * 60
FIJI_STARTUP_TIMEOUT_SECONDS = 90

# How long the output file must remain unchanged before it is considered stable.
CSV_STABLE_SECONDS = 4.0

# Keep Fiji open briefly after a successful output so its log can be saved.
POST_RUN_GRACE_SECONDS = 2.0


# =============================================================================
# 2. FILOPODYAN PARAMETERS
# =============================================================================
#
# These are the parameters from your existing pipeline.
#
# C1 = Target channel
# C2 = Lifeact / mapping channel
#
# Filopodyan's source calls these fields mapC and measureC.
#
# IMPORTANT FIX FROM THE PREVIOUS VERSION:
# Filopodyan's GUI converts a positive "base back frames" entry into a
# NEGATIVE internal backFrames value. Therefore 3 must be passed to the plugin
# as -3, not +3.
#

MAP_CHANNEL = 2
MEASURE_CHANNEL = 1

THRESHOLD = "Triangle"
ADAPTIVE = False
FIT_TIP_TO_MEASURE_C = False
JOIN_FRAGMENTS = False

ED_ITERATIONS = 8
BASE_BACK_FRAMES = 3
LOG_SIGMA = 4.0

# Track filtering
MIN_START_FRAME = 1
MIN_FRAMES = 5
MIN_MAX_LENGTH_UM = 0.5
MIN_LENGTH_CHANGE_UM = 2.0
MIN_MAX_DCTM_UM = 0.0
MIN_MAX_DCBM_UM = 0.0
MAX_MEAN_WAVINESS = 1.0


# =============================================================================
# 3. OPTIONAL DIAGNOSTIC BEHAVIOR
# =============================================================================
#
# The analysis itself always uses the parameters above.
#
# If Filopodyan produces a CSV containing only "dT", that is NOT treated as a
# successful filopodia dataset. The script records NO_TRACKS and preserves the
# Fiji/Filopodyan log so the problem can be diagnosed without silently making
# up data or changing the biological thresholds.
#
SAVE_FIJI_LOG = True


# =============================================================================
# 4. DIRECTORIES
# =============================================================================

for directory in (OUTPUT_DIR, TABLE_DIR, THUMBNAIL_DIR, LOG_DIR):
    os.makedirs(directory, exist_ok=True)


# =============================================================================
# 5. HELPERS
# =============================================================================

def java_string(value: str) -> str:
    """Escape a Python string for insertion into a BeanShell script."""
    return (
        str(value)
        .replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
    )


def filopodyan_sanitized_title(filename: str) -> str:
    """Match Filopodyan_.java: replace non-alphanumeric/underscore chars."""
    return re.sub(r"[^a-zA-Z0-9_]", "_", filename)


def expected_csv(filename: str, kind: str = "Filopodia") -> str:
    safe = filopodyan_sanitized_title(filename)
    return os.path.join(TABLE_DIR, f"{safe}_{kind}.csv")


def expected_log(filename: str) -> str:
    safe = filopodyan_sanitized_title(filename)
    return os.path.join(LOG_DIR, f"{safe}_FilopodyanLog.txt")


def remove_if_exists(path: str) -> None:
    try:
        if os.path.isfile(path):
            os.remove(path)
    except OSError as exc:
        print(f"[WARNING] Could not remove {path}: {exc}")


# =============================================================================
# 6. VISUALIZATION
# =============================================================================

def create_dual_channel_overlay(frame_c1, frame_c2):
    c1 = np.asarray(frame_c1, dtype=np.float32)
    c2 = np.asarray(frame_c2, dtype=np.float32)

    def normalize(arr):
        low, high = np.percentile(arr, (1, 99))
        if high <= low:
            return np.zeros(arr.shape, dtype=np.uint8)
        return np.clip(
            (arr - low) * 255.0 / (high - low), 0, 255
        ).astype(np.uint8)

    overlay = np.zeros((c1.shape[0], c1.shape[1], 3), dtype=np.uint8)

    # Red = mapping / Lifeact C2
    overlay[:, :, 0] = normalize(c2)
    # Green = measured Target C1
    overlay[:, :, 1] = normalize(c1)

    return Image.fromarray(overlay)


def make_snapshot_crops(tiff_path: str, filename: str):
    """Read only one timepoint for the report; do not load the whole movie."""
    crops = []

    with tifffile.TiffFile(tiff_path) as tif:
        series = tif.series[0]
        axes = series.axes

        if axes != "TCYX":
            raise ValueError(
                f"{filename}: expected TCYX TIFF hyperstack for report snapshots; "
                f"found axes={axes}, shape={series.shape}"
            )

        n_t, n_c, height, width = series.shape
        if n_c < max(MAP_CHANNEL, MEASURE_CHANNEL):
            raise ValueError(
                f"{filename}: only {n_c} channel(s); expected at least 2."
            )

        t = n_t // 2

        c1 = series.pages[t * n_c + (MEASURE_CHANNEL - 1)].asarray()
        c2 = series.pages[t * n_c + (MAP_CHANNEL - 1)].asarray()
        overlay = create_dual_channel_overlay(c1, c2)

        threshold = np.percentile(c2, 99.5)
        y, x = np.where(c2 >= threshold)

        if len(x) == 0:
            centers = [(width // 2, height // 2)]
        else:
            centers = []
            grid = 2
            for gy in range(grid):
                for gx in range(grid):
                    x0 = gx * width // grid
                    x1 = (gx + 1) * width // grid
                    y0 = gy * height // grid
                    y1 = (gy + 1) * height // grid
                    keep = (
                        (x >= x0) & (x < x1) &
                        (y >= y0) & (y < y1)
                    )
                    if np.any(keep):
                        xx = x[keep]
                        yy = y[keep]
                        best = np.argmax(c2[yy, xx])
                        centers.append((int(xx[best]), int(yy[best])))

        for i, (cx, cy) in enumerate(centers[:4]):
            box = (
                max(0, cx - 75),
                max(0, cy - 75),
                min(width, cx + 75),
                min(height, cy + 75),
            )
            crop = overlay.crop(box)
            path = os.path.join(
                THUMBNAIL_DIR,
                f"{filopodyan_sanitized_title(filename)}_T{t + 1}_crop{i + 1}.jpg",
            )
            crop.save(path, "JPEG", quality=95)
            crops.append(path)

    return crops, t + 1


# =============================================================================
# 7. FIJI / BEANSHELL SCRIPT
# =============================================================================
#
# Fiji is launched ONCE for the complete TIFF batch. Each movie is opened,
# analyzed, and closed inside the same Fiji/JVM instance.
#
# The existing Fiji installation and existing Filopodyan classes are used.
# Nothing is installed or downloaded by this script.
# =============================================================================

def create_fiji_batch_script(tiff_files: list[str]) -> str:
    output_dir = java_string(OUTPUT_DIR)
    table_dir = java_string(TABLE_DIR)
    status_file = java_string(
        os.path.join(OUTPUT_DIR, "_filopodyan_status.txt")
    )

    java_files = ",\n".join(
        f'    "{java_string(path)}"' for path in tiff_files
    )

    script = r"""
import ij.IJ;
import ij.ImagePlus;
import java.io.File;
import java.io.FileWriter;
import java.lang.reflect.Field;
import java.util.concurrent.atomic.AtomicInteger;
import javax.swing.JTabbedPane;
import javax.swing.JScrollPane;
import javax.swing.JTextPane;
import javax.swing.text.Document;

String outputDir = "__OUTPUT_DIR__";
String tableDir = "__TABLE_DIR__";
String statusPath = "__STATUS_FILE__";

String[] inputFiles = new String[] {
__INPUT_FILES__
};

new File(outputDir).mkdirs();
new File(tableDir).mkdirs();

void status(String text) {
    try {
        FileWriter fw = new FileWriter(statusPath, false);
        fw.write(text + "\n");
        fw.close();
    } catch (Exception e) {
        IJ.log("[AUTOMATION] Could not write status: " + e);
    }
}

void setInt(Object obj, String name, int value) {
    try {
        Field f = obj.getClass().getField(name);
        f.setInt(obj, value);
    } catch (Exception e) {
        throw new RuntimeException("Could not set int field " + name + ": " + e);
    }
}

void setDouble(Object obj, String name, double value) {
    try {
        Field f = obj.getClass().getField(name);
        f.setDouble(obj, value);
    } catch (Exception e) {
        throw new RuntimeException("Could not set double field " + name + ": " + e);
    }
}

void setBoolean(Object obj, String name, boolean value) {
    try {
        Field f = obj.getClass().getField(name);
        f.setBoolean(obj, value);
    } catch (Exception e) {
        throw new RuntimeException("Could not set boolean field " + name + ": " + e);
    }
}

void setString(Object obj, String name, String value) {
    try {
        Field f = obj.getClass().getField(name);
        f.set(obj, value);
    } catch (Exception e) {
        throw new RuntimeException("Could not set String field " + name + ": " + e);
    }
}

void saveFilopodyanLog(BatchFilopodyan bgui, String path) {
    try {
        if (bgui.log == null) return;

        JTabbedPane pane = bgui.log.getPane();
        StringBuilder all = new StringBuilder();
        int n = pane.getTabCount();

        for (int i = 0; i < n; i++) {
            all.append("===== ")
               .append(pane.getTitleAt(i))
               .append(" =====\n");

            JScrollPane scroll =
                (JScrollPane)pane.getComponentAt(i);

            JTextPane text =
                (JTextPane)scroll.getViewport().getView();

            Document doc = text.getDocument();

            all.append(
                doc.getText(0, doc.getLength())
            ).append("\n\n");
        }

        IJ.saveString(all.toString(), path);

    } catch (Exception e) {
        IJ.log("[AUTOMATION] Could not save Filopodyan log: " + e);
    }
}

void configureBatch(BatchFilopodyan bgui) {

    setInt(bgui, "mapC", __MAP_CHANNEL__);
    setInt(bgui, "measureC", __MEASURE_CHANNEL__);
    setString(bgui, "threshold", "__THRESHOLD__");
    setBoolean(bgui, "adaptive", __ADAPTIVE__);
    setBoolean(bgui, "fit", __FIT__);
    setBoolean(bgui, "join", __JOIN__);
    setInt(bgui, "eds", __ED_ITERATIONS__);
    setInt(bgui, "backFrames", __BACK_FRAMES_INTERNAL__);
    setDouble(bgui, "sigma", __LOG_SIGMA__);

    setBoolean(bgui, "filoTable", true);
    setBoolean(bgui, "coordTable", true);
    setBoolean(bgui, "bodyTable", true);

    setBoolean(bgui, "boundaryAnalysis", false);
    setBoolean(bgui, "kymographs", false);
    setBoolean(bgui, "ccf", false);
    setBoolean(bgui, "basePlot", false);
    setBoolean(bgui, "tipPlot", false);
    setBoolean(bgui, "time", false);
    setBoolean(bgui, "showBackground", false);
    setBoolean(bgui, "verbose", true);

    setString(bgui, "path", outputDir);
    setInt(bgui, "start", __MIN_START_FRAME__);
    setInt(bgui, "frames", __MIN_FRAMES__);
    setDouble(bgui, "length", __MIN_MAX_LENGTH_UM__);
    setDouble(bgui, "dl", __MIN_LENGTH_CHANGE_UM__);
    setDouble(bgui, "dctm", __MIN_MAX_DCTM_UM__);
    setDouble(bgui, "dcbm", __MIN_MAX_DCBM_UM__);
    setDouble(bgui, "waviness", __MAX_MEAN_WAVINESS__);

    bgui.count = new AtomicInteger(1);
    bgui.log = FilopodyanLog.get();
}

void deleteOldOutputs(String tableDir, String safeTitle) {
    File filoCsv =
        new File(tableDir + File.separator + safeTitle + "_Filopodia.csv");

    File coordCsv =
        new File(tableDir + File.separator + safeTitle + "_Coordinates.csv");

    File bodyCsv =
        new File(tableDir + File.separator + safeTitle + "_Bodies.csv");

    if (filoCsv.exists()) filoCsv.delete();
    if (coordCsv.exists()) coordCsv.delete();
    if (bodyCsv.exists()) bodyCsv.delete();
}

status("FIJI_STARTED");

IJ.log("============================================================");
IJ.log("[AUTOMATION] Single-Fiji Filopodyan batch started.");
IJ.log("[AUTOMATION] Number of TIFF files: " + inputFiles.length);

try {
    Class.forName("BatchFilopodyan");
    Class.forName("Filopodyan_");
    Class.forName("FilopodyanProcessor");

    IJ.log("[Filopodyan] Existing Filopodyan classes found in this Fiji.");
    status("PLUGIN_FOUND");

} catch (Exception e) {
    status("PLUGIN_ERROR|" + e.toString());
    throw new RuntimeException(
        "Installed Filopodyan classes not found: " + e
    );
}

for (int fileIndex = 0; fileIndex < inputFiles.length; fileIndex++) {

    String inputPath = inputFiles[fileIndex];
    File inputFile = new File(inputPath);

    String safeTitle =
        inputFile.getName().replaceAll("[^a-zA-Z0-9_]", "_");

    File filoCsv =
        new File(
            tableDir + File.separator +
            safeTitle + "_Filopodia.csv"
        );

    String logPath =
        outputDir + File.separator +
        "logs" + File.separator +
        safeTitle + "_FilopodyanLog.txt";

    new File(outputDir + File.separator + "logs").mkdirs();

    IJ.log("");
    IJ.log("============================================================");
    IJ.log(
        "[Filopodyan] FILE " +
        (fileIndex + 1) + "/" + inputFiles.length +
        ": " + inputFile.getName()
    );
    IJ.log("============================================================");

    status(
        "STARTING|" +
        (fileIndex + 1) + "/" + inputFiles.length +
        "|" + inputPath
    );

    deleteOldOutputs(tableDir, safeTitle);

    ImagePlus image = null;
    BatchFilopodyan bgui = null;

    try {

        IJ.log("[Filopodyan] Opening: " + inputPath);

        image = IJ.openImage(inputPath);

        if (image == null) {
            throw new RuntimeException(
                "Could not open image: " + inputPath
            );
        }

        IJ.log(
            "[Filopodyan] Image dimensions: C=" +
            image.getNChannels() +
            " Z=" + image.getNSlices() +
            " T=" + image.getNFrames() +
            " W=" + image.getWidth() +
            " H=" + image.getHeight()
        );

        if (image.getNChannels() < 2) {
            throw new RuntimeException("Expected at least 2 channels.");
        }

        status(
            "IMAGE_OPENED|" +
            (fileIndex + 1) + "/" + inputFiles.length +
            "|" + inputPath
        );

        // IMPORTANT FIX: construct before setting fields.
        bgui = new BatchFilopodyan();
        configureBatch(bgui);

        IJ.log("[Filopodyan] Map C = " + bgui.mapC);
        IJ.log("[Filopodyan] Measure C = " + bgui.measureC);
        IJ.log("[Filopodyan] Threshold = " + bgui.threshold);
        IJ.log("[Filopodyan] Adaptive = " + bgui.adaptive);
        IJ.log("[Filopodyan] Fit tip = " + bgui.fit);
        IJ.log("[Filopodyan] Join = " + bgui.join);
        IJ.log("[Filopodyan] ED iterations = " + bgui.eds);
        IJ.log("[Filopodyan] Internal backFrames = " + bgui.backFrames);
        IJ.log("[Filopodyan] LoG sigma = " + bgui.sigma);

        Filopodyan_ bounder = new Filopodyan_();
        bounder.setImp(image);
        bounder.batch = true;
        bounder.bgui = bgui;

        IJ.log("[Filopodyan] BEFORE bounder.filopodia(false)");

        status(
            "RUNNING|" +
            (fileIndex + 1) + "/" + inputFiles.length +
            "|" + inputPath
        );

        bounder.filopodia(false);

        IJ.log("[Filopodyan] AFTER bounder.filopodia(false)");

        status(
            "FILOPODYA_RETURNED|" +
            (fileIndex + 1) + "/" + inputFiles.length +
            "|" + inputPath
        );

        long waitStart = System.currentTimeMillis();

        while (bgui.count.get() > 0) {

            Thread.sleep(500);

            if (
                System.currentTimeMillis() - waitStart >
                __TIMEOUT_MS__
            ) {
                IJ.log("[Filopodyan] ERROR: output worker timeout.");

                status(
                    "OUTPUT_TIMEOUT|" +
                    (fileIndex + 1) + "/" + inputFiles.length +
                    "|" + inputPath
                );

                break;
            }
        }

        saveFilopodyanLog(bgui, logPath);

        if (filoCsv.exists()) {

            IJ.log(
                "[Filopodyan] Filopodia CSV exists: " +
                filoCsv.getAbsolutePath()
            );

            status(
                "CSV_EXISTS|" +
                (fileIndex + 1) + "/" + inputFiles.length +
                "|" + filoCsv.getAbsolutePath()
            );

        } else {

            IJ.log("[Filopodyan] ERROR: Filopodia CSV was not created.");

            status(
                "CSV_MISSING|" +
                (fileIndex + 1) + "/" + inputFiles.length +
                "|" + filoCsv.getAbsolutePath()
            );
        }

    } catch (Exception e) {

        IJ.log(
            "[Filopodyan] ERROR processing " +
            inputPath + ": " + e
        );

        status(
            "FILE_ERROR|" +
            (fileIndex + 1) + "/" + inputFiles.length +
            "|" + e.toString()
        );

    } finally {

        if (image != null) {
            try {
                image.changes = false;
                image.close();
            } catch (Exception ignored) {}
        }

        IJ.log(
            "[Filopodyan] Finished file " +
            (fileIndex + 1) + "/" + inputFiles.length
        );
    }
}

status("BATCH_FINISHED");

IJ.log("");
IJ.log("============================================================");
IJ.log("[AUTOMATION] COMPLETE: all TIFF files processed.");
IJ.log("============================================================");
"""

    replacements = {
        "__OUTPUT_DIR__": output_dir,
        "__TABLE_DIR__": table_dir,
        "__STATUS_FILE__": status_file,
        "__INPUT_FILES__": java_files,
        "__MAP_CHANNEL__": str(MAP_CHANNEL),
        "__MEASURE_CHANNEL__": str(MEASURE_CHANNEL),
        "__THRESHOLD__": java_string(THRESHOLD),
        "__ADAPTIVE__": str(ADAPTIVE).lower(),
        "__FIT__": str(FIT_TIP_TO_MEASURE_C).lower(),
        "__JOIN__": str(JOIN_FRAGMENTS).lower(),
        "__ED_ITERATIONS__": str(ED_ITERATIONS),
        "__BACK_FRAMES_INTERNAL__": str(-abs(BASE_BACK_FRAMES)),
        "__LOG_SIGMA__": str(LOG_SIGMA),
        "__MIN_START_FRAME__": str(MIN_START_FRAME),
        "__MIN_FRAMES__": str(MIN_FRAMES),
        "__MIN_MAX_LENGTH_UM__": str(MIN_MAX_LENGTH_UM),
        "__MIN_LENGTH_CHANGE_UM__": str(MIN_LENGTH_CHANGE_UM),
        "__MIN_MAX_DCTM_UM__": str(MIN_MAX_DCTM_UM),
        "__MIN_MAX_DCBM_UM__": str(MIN_MAX_DCBM_UM),
        "__MAX_MEAN_WAVINESS__": str(MAX_MEAN_WAVINESS),
        "__TIMEOUT_MS__": str(FILOPODYAN_TIMEOUT_SECONDS * 1000),
    }

    for key, value in replacements.items():
        script = script.replace(key, value)

    return textwrap.dedent(script)


# =============================================================================
# 8. FIJI PROCESS CONTROL — SINGLE FIJI INSTANCE
# =============================================================================

def terminate_fiji(process):
    if process is None or process.poll() is not None:
        return

    try:
        process.terminate()
        process.wait(timeout=10)
        return
    except Exception:
        pass

    try:
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        pass


def read_status():
    path = os.path.join(OUTPUT_DIR, "_filopodyan_status.txt")

    if not os.path.isfile(path):
        return "NO_STATUS_FILE"

    try:
        with open(
            path,
            "r",
            encoding="utf-8",
            errors="replace",
        ) as f:
            lines = [line.strip() for line in f if line.strip()]

        return lines[-1] if lines else "EMPTY_STATUS_FILE"

    except Exception as exc:
        return f"STATUS_READ_ERROR|{type(exc).__name__}|{exc}"


def wait_for_batch_completion(process, timeout_seconds: int):
    start = time.time()
    last_status = None

    while time.time() - start < timeout_seconds:

        status_text = read_status()

        if status_text != last_status:
            last_status = status_text
            print(f"[FIJI] Status: {status_text}")

        if status_text == "BATCH_FINISHED":
            return True

        if process.poll() is not None:
            print(
                "[FIJI] Fiji process exited before batch completion. "
                f"exit={process.poll()}"
            )
            return False

        time.sleep(1)

    print(
        "[FIJI] ERROR: batch timeout after "
        f"{timeout_seconds // 60} minutes."
    )
    return False


def run_filopodyan(tiff_files):

    if not os.path.isfile(FIJI_EXE):
        raise FileNotFoundError(
            f"Fiji executable not found:\n{FIJI_EXE}"
        )

    script_path = os.path.join(
        SCRIPT_DIR,
        "_automated_filopodyan_batch.bsh",
    )

    status_path = os.path.join(
        OUTPUT_DIR,
        "_filopodyan_status.txt",
    )

    remove_if_exists(status_path)

    with open(script_path, "w", encoding="utf-8") as f:
        f.write(create_fiji_batch_script(tiff_files))

    print()
    print("[INFO] Using the EXISTING Fiji installation only.")
    print(f"[INFO] Fiji: {FIJI_EXE}")
    print("[INFO] Using the EXISTING Filopodyan_ plugin classes already in Fiji.")
    print(f"[INFO] Movies: {len(tiff_files)}")
    print(
        "[INFO] Fiji processes used: 1 "
        "(all movies run sequentially inside the same Fiji instance)"
    )

    batch_timeout = FILOPODYAN_TIMEOUT_SECONDS * len(tiff_files)

    print(
        "[INFO] Complete-batch timeout: "
        f"{batch_timeout // 60} minutes"
    )

    command = [
        FIJI_EXE,
        "--console",
        "--run",
        script_path,
    ]

    print()
    print("[FIJI] Launching ONE Fiji process for the entire batch...")
    print(f"[FIJI] Command: {' '.join(command)}")

    process = None

    try:
        process = subprocess.Popen(
            command,
            cwd=os.path.dirname(FIJI_EXE),
            stdout=None,
            stderr=None,
        )

    except Exception as exc:

        print(
            "[FIJI] ERROR launching Fiji: "
            f"{type(exc).__name__}: {exc}"
        )

        return [
            {
                "filename": os.path.basename(path),
                "status": "FIJI_LAUNCH_ERROR",
                "csv": expected_csv(
                    os.path.basename(path),
                    "Filopodia",
                ),
                "log": expected_log(os.path.basename(path)),
                "message": f"{type(exc).__name__}: {exc}",
            }
            for path in tiff_files
        ]

    print(f"[FIJI] PID = {process.pid}")

    print(
        "[FIJI] Waiting for Fiji/BeanShell startup "
        f"(up to {FIJI_STARTUP_TIMEOUT_SECONDS}s)..."
    )

    startup_deadline = time.time() + FIJI_STARTUP_TIMEOUT_SECONDS
    startup_seen = False

    while time.time() < startup_deadline:

        status_text = read_status()

        if status_text != "NO_STATUS_FILE":
            startup_seen = True
            print(f"[FIJI] Startup status: {status_text}")
            break

        if process.poll() is not None:
            break

        time.sleep(1)

    if not startup_seen:

        message = (
            "Fiji did not reach the BeanShell startup checkpoint "
            f"within {FIJI_STARTUP_TIMEOUT_SECONDS} seconds. "
            f"Process exit={process.poll()}"
        )

        print(f"[FIJI] ERROR: {message}")
        terminate_fiji(process)

        return [
            {
                "filename": os.path.basename(path),
                "status": "FIJI_STARTUP_TIMEOUT",
                "csv": expected_csv(
                    os.path.basename(path),
                    "Filopodia",
                ),
                "log": expected_log(os.path.basename(path)),
                "message": message,
            }
            for path in tiff_files
        ]

    print(
        "[FIJI] Fiji is running. "
        "Processing all movies in this same instance..."
    )

    completed = wait_for_batch_completion(
        process,
        batch_timeout,
    )

    if completed:
        print("[FIJI] Batch completion checkpoint received.")
    else:
        print("[FIJI] Batch did not reach normal completion.")

    print("[FIJI] Closing the single Fiji process...")
    terminate_fiji(process)
    time.sleep(2)

    results = []

    for tiff_path in tiff_files:

        filename = os.path.basename(tiff_path)
        csv_path = expected_csv(filename, "Filopodia")

        if os.path.isfile(csv_path):
            results.append(
                {
                    "filename": filename,
                    "status": "CSV_CREATED",
                    "csv": csv_path,
                    "log": expected_log(filename),
                    "message": "BATCH_FINISHED",
                }
            )
        else:
            results.append(
                {
                    "filename": filename,
                    "status": "CSV_FAILED",
                    "csv": csv_path,
                    "log": expected_log(filename),
                    "message": "No completed Filopodyan CSV was found.",
                }
            )

    return results


# =============================================================================
# 9. CSV VALIDATION / PARSING
# =============================================================================

def inspect_filopodyan_csv(csv_path: str):
    df = pd.read_csv(csv_path)
    columns = [str(c) for c in df.columns]

    track_time_columns = [
        c for c in columns if re.fullmatch(r"T \(\d+\)", c)
    ]
    track_length_columns = [
        c for c in columns if re.fullmatch(r"Length \(\d+\)", c)
    ]

    # A dT-only Filopodyan output is a valid file produced by the plugin, but
    # it contains zero usable track columns and therefore is NOT a successful
    # filopodia dataset.
    if not track_time_columns:
        if columns == ["dT"]:
            return {
                "status": "NO_TRACKS",
                "dataframe": df,
                "track_count": 0,
                "message": "Filopodyan produced dT only; no filopodia track columns were written.",
            }

        return {
            "status": "NO_TRACK_COLUMNS",
            "dataframe": df,
            "track_count": 0,
            "message": f"No T (n) columns found. Columns: {columns}",
        }

    return {
        "status": "TRACKS_PRESENT",
        "dataframe": df,
        "track_count": len(track_time_columns),
        "time_columns": track_time_columns,
        "length_columns": track_length_columns,
        "message": "Track columns found.",
    }


def parse_filopodyan_csv(csv_path: str):
    info = inspect_filopodyan_csv(csv_path)
    if info["status"] != "TRACKS_PRESENT":
        return info

    df = info["dataframe"]
    time_columns = info["time_columns"]
    length_columns = info["length_columns"]

    rows = []

    for row_index, row in df.iterrows():
        active_tracks = 0
        new_formations = 0
        lengths = []

        # dT is relative output time. Actual T(n) values are absolute plugin
        # frame values. For formation counting, compare to the row's absolute
        # frame using dT + internal backFrames.
        if "dT" in df.columns:
            try:
                actual_frame = float(row["dT"]) - (-abs(BASE_BACK_FRAMES))
            except Exception:
                actual_frame = float(row_index + 1)
        else:
            actual_frame = float(row_index + 1)

        for col in time_columns:
            value = row[col]
            if pd.notna(value):
                active_tracks += 1
                try:
                    if int(round(float(value))) == int(round(actual_frame)):
                        new_formations += 1
                except Exception:
                    pass

        for col in length_columns:
            value = row[col]
            if pd.notna(value):
                try:
                    lengths.append(float(value))
                except Exception:
                    pass

        rows.append(
            {
                "Timepoint": f"T_{row_index + 1}",
                "Protrusion Count": active_tracks,
                "Mean Length (µm)": float(np.mean(lengths)) if lengths else np.nan,
                "New Formations": new_formations,
            }
        )

    return {
        **info,
        "status": "TRACKS_PRESENT",
        "parsed": pd.DataFrame(rows),
    }


# =============================================================================
# 10. REPORT HELPERS
# =============================================================================

def add_text(slide, text, left, top, width, height, size=14, bold=False):
    box = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    p = box.text_frame.paragraphs[0]
    p.text = str(text)
    p.font.size = Pt(size)
    p.font.bold = bold
    return box


def add_table_slide(prs, filename, df, blank):
    chunk_size = 25
    for start in range(0, len(df), chunk_size):
        chunk = df.iloc[start:start + chunk_size]
        slide = prs.slides.add_slide(blank)

        add_text(
            slide,
            f"{filename} — Filopodyan Results (T{start + 1}–T{start + len(chunk)})",
            0.5, 0.35, 12.3, 0.55, size=20, bold=True,
        )

        rows = len(chunk) + 1
        columns = 4
        table_shape = slide.shapes.add_table(
            rows, columns,
            Inches(0.7), Inches(1.25), Inches(11.9),
            Inches(min(5.8, 0.22 * rows)),
        )
        table = table_shape.table

        headers = [
            "Timepoint",
            "Protrusion Count",
            "Mean Length (µm)",
            "New Formations",
        ]

        for c, header in enumerate(headers):
            table.cell(0, c).text = header
            for p in table.cell(0, c).text_frame.paragraphs:
                p.font.bold = True
                p.font.size = Pt(10)

        for r, (_, data) in enumerate(chunk.iterrows(), start=1):
            table.cell(r, 0).text = str(data["Timepoint"])
            table.cell(r, 1).text = str(int(data["Protrusion Count"]))
            length = data["Mean Length (µm)"]
            table.cell(r, 2).text = "N/A" if pd.isna(length) else f"{length:.3f}"
            table.cell(r, 3).text = str(int(data["New Formations"]))
            for c in range(columns):
                for p in table.cell(r, c).text_frame.paragraphs:
                    p.font.size = Pt(9)


def add_no_track_slide(prs, filename, info, log_path, blank):
    slide = prs.slides.add_slide(blank)
    add_text(
        slide,
        f"Filopodyan Diagnostic: {filename}",
        0.6, 0.7, 12.0, 0.7, size=24, bold=True,
    )
    add_text(
        slide,
        "NO FILOPODIA TRACK COLUMNS WERE WRITTEN",
        0.6, 1.7, 12.0, 0.6, size=18, bold=True,
    )
    add_text(
        slide,
        info["message"],
        0.6, 2.5, 12.0, 0.7, size=14,
    )
    add_text(
        slide,
        f"CSV: {expected_csv(filename, 'Filopodia')}",
        0.6, 3.4, 12.0, 0.45, size=11,
    )
    add_text(
        slide,
        f"Filopodyan log: {log_path}",
        0.6, 4.0, 12.0, 0.45, size=11,
    )
    add_text(
        slide,
        (
            "This report intentionally does not calculate filopodia counts or lengths "
            "from a dT-only table. The plugin output is preserved for diagnosis."
        ),
        0.6, 4.8, 12.0, 1.0, size=13,
    )


def build_report(dataset_results, output_path):
    print()
    print("[BUILD] Creating PowerPoint report...")

    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Title
    slide = prs.slides.add_slide(blank)
    add_text(
        slide,
        "Automated Filopodyan Analysis",
        0.7, 2.3, 12.0, 0.7, size=30, bold=True,
    )
    add_text(
        slide,
        "Existing Fiji + installed Filopodyan two-channel time-lapse pipeline",
        0.7, 3.2, 12.0, 0.5, size=16,
    )
    add_text(
        slide,
        f"Map C{MAP_CHANNEL} = Lifeact | Measure C{MEASURE_CHANNEL} = Target C",
        0.7, 4.0, 12.0, 0.45, size=13, bold=True,
    )

    for filename, result in dataset_results.items():
        info = result["info"]

        if info["status"] != "TRACKS_PRESENT":
            add_no_track_slide(
                prs, filename, info, result["log"], blank
            )
            continue

        df = info["parsed"]
        crops = result.get("crops", [])

        slide = prs.slides.add_slide(blank)
        add_text(
            slide,
            f"Filopodyan Quantitative Analysis: {filename}",
            0.5, 0.35, 12.3, 0.55, size=22, bold=True,
        )
        add_text(
            slide,
            (
                f"Map C{MAP_CHANNEL} = Lifeact | "
                f"Measure C{MEASURE_CHANNEL} = Target C | "
                f"Threshold = {THRESHOLD} | LoG σ = {LOG_SIGMA}"
            ),
            0.5, 0.95, 12.3, 0.4, size=11, bold=True,
        )

        maximum_count = int(df["Protrusion Count"].max())
        mean_length = df["Mean Length (µm)"].mean()
        total_new = int(df["New Formations"].sum())

        add_text(slide, f"Maximum active filopodia: {maximum_count}",
                 0.6, 1.5, 3.8, 0.45, size=16, bold=True)
        add_text(slide, f"Mean measured length: {mean_length:.3f} µm",
                 4.5, 1.5, 4.0, 0.45, size=16, bold=True)
        add_text(slide, f"Detected formation events: {total_new}",
                 8.7, 1.5, 3.8, 0.45, size=16, bold=True)

        add_text(slide, "Representative fluorescence snapshots",
                 7.3, 2.15, 5.2, 0.4, size=13, bold=True)

        for i, image_path in enumerate(crops):
            col = i % 2
            row = i // 2
            slide.shapes.add_picture(
                image_path,
                Inches(7.3 + col * 2.8),
                Inches(2.65 + row * 2.0),
                width=Inches(2.55),
                height=Inches(1.7),
            )

        add_table_slide(prs, filename, df, blank)

    prs.save(output_path)
    print(f"[SUCCESS] PowerPoint saved: {output_path}")


# =============================================================================
# 11. SUMMARY CSV
# =============================================================================

def write_run_summary(rows):
    pd.DataFrame(rows).to_csv(OUTPUT_SUMMARY_CSV, index=False)
    print(f"[INFO] Run summary saved: {OUTPUT_SUMMARY_CSV}")


# =============================================================================
# 12. MAIN
# =============================================================================

def main():
    print("=" * 78)
    print("AUTOMATED FIJI + FILOPODYAN FILOPODIA ANALYSIS")
    print("=" * 78)
    print("[INFO] No Fiji or Filopodyan installation is performed by this script.")
    print("[INFO] The existing installed Fiji and Filopodyan classes are used.")
    print()
    print("[PARAMETERS]")
    print(f"  Map C: {MAP_CHANNEL}")
    print(f"  Measure C: {MEASURE_CHANNEL}")
    print(f"  Threshold: {THRESHOLD}")
    print(f"  Adaptive: {ADAPTIVE}")
    print(f"  Fit tip: {FIT_TIP_TO_MEASURE_C}")
    print(f"  Join fragments: {JOIN_FRAGMENTS}")
    print(f"  ED iterations: {ED_ITERATIONS}")
    print(f"  Base back frames entered: {BASE_BACK_FRAMES}")
    print(f"  Internal Filopodyan backFrames: {-abs(BASE_BACK_FRAMES)}")
    print(f"  LoG sigma: {LOG_SIGMA}")
    print(f"  Min start frame: {MIN_START_FRAME}")
    print(f"  Min frames: {MIN_FRAMES}")
    print(f"  Min max length: {MIN_MAX_LENGTH_UM} µm")
    print(f"  Min length change: {MIN_LENGTH_CHANGE_UM} µm")
    print(f"  Min max DCTM: {MIN_MAX_DCTM_UM} µm")
    print(f"  Min max DCBM: {MIN_MAX_DCBM_UM} µm")
    print(f"  Max mean waviness: {MAX_MEAN_WAVINESS}")

    if not os.path.isdir(IMAGE_FOLDER):
        raise FileNotFoundError(
            f"Input directory does not exist:\n{IMAGE_FOLDER}"
        )

    if not os.path.isfile(FIJI_EXE):
        raise FileNotFoundError(
            f"Fiji executable does not exist:\n{FIJI_EXE}"
        )

    tiff_files = [
        os.path.join(IMAGE_FOLDER, name)
        for name in sorted(os.listdir(IMAGE_FOLDER))
        if name.lower().endswith((".tif", ".tiff"))
    ]

    if not tiff_files:
        raise RuntimeError("No TIFF files found in the input folder.")

    print(f"\n[INFO] Found {len(tiff_files)} TIFF file(s).")

    run_results = run_filopodyan(tiff_files)
    summary_rows = []
    dataset_results = {}

    for run in run_results:
        filename = run["filename"]
        csv_path = run["csv"]
        log_path = run["log"]

        row = {
            "Filename": filename,
            "RunStatus": run["status"],
            "CSV": csv_path,
            "FilopodyanLog": log_path,
            "TrackStatus": "NOT_READ",
            "TrackCount": 0,
            "Message": run["message"],
        }

        if os.path.isfile(csv_path):
            try:
                info = parse_filopodyan_csv(csv_path)
                row["TrackStatus"] = info["status"]
                row["TrackCount"] = info.get("track_count", 0)
                row["Message"] = info["message"]

                dataset_results[filename] = {
                    "info": info,
                    "log": log_path,
                    "csv": csv_path,
                }

                if info["status"] == "TRACKS_PRESENT":
                    try:
                        tiff_path = os.path.join(IMAGE_FOLDER, filename)
                        crops, middle_frame = make_snapshot_crops(
                            tiff_path, filename
                        )
                        dataset_results[filename]["crops"] = crops
                        dataset_results[filename]["middle_frame"] = middle_frame
                    except Exception as exc:
                        print(
                            f"[WARNING] Could not create report snapshots for {filename}: "
                            f"{type(exc).__name__}: {exc}"
                        )
                        dataset_results[filename]["crops"] = []
            except Exception as exc:
                row["TrackStatus"] = "CSV_PARSE_ERROR"
                row["Message"] = f"{type(exc).__name__}: {exc}"
        else:
            print(f"[READ] CSV missing for {filename}")

        summary_rows.append(row)

    write_run_summary(summary_rows)

    print()
    print("=" * 78)
    print("RESULT SUMMARY")
    print("=" * 78)

    track_present = 0
    no_tracks = 0
    failed = 0

    for row in summary_rows:
        print(
            f"{row['Filename']}: "
            f"run={row['RunStatus']} | "
            f"tracks={row['TrackStatus']} | "
            f"count={row['TrackCount']}"
        )
        if row["TrackStatus"] == "TRACKS_PRESENT":
            track_present += 1
        elif row["TrackStatus"] in ("NO_TRACKS", "NO_TRACK_COLUMNS"):
            no_tracks += 1
        else:
            failed += 1

    print()
    print(f"[INFO] Files with usable track columns: {track_present}/{len(tiff_files)}")
    print(f"[INFO] Files with no track columns: {no_tracks}/{len(tiff_files)}")
    print(f"[INFO] Other failures/errors: {failed}/{len(tiff_files)}")

    # Always create a report if at least one CSV exists. This makes the dT-only
    # situation visible rather than hiding it behind a Python parser exception.
    if dataset_results:
        build_report(dataset_results, OUTPUT_PPTX)
    else:
        print("[WARNING] No Filopodyan CSVs were available; no PowerPoint created.")

    print()
    print("=" * 78)
    print("PIPELINE COMPLETE")
    print("=" * 78)
    print(f"[INFO] Output folder: {OUTPUT_DIR}")
    print(f"[INFO] Tables: {TABLE_DIR}")
    print(f"[INFO] Logs: {LOG_DIR}")
    print(f"[INFO] Summary: {OUTPUT_SUMMARY_CSV}")
    if os.path.isfile(OUTPUT_PPTX):
        print(f"[INFO] PowerPoint: {OUTPUT_PPTX}")


if __name__ == "__main__":
    main()
