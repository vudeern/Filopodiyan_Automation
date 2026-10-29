from PIL import Image
import os

FILE_PATH = r"C:\Users\udeer\Downloads\Filopodia analysis\Filopodia analysis\MOV_9.tif"


print("=" * 60)
print("TIFF STRUCTURE INSPECTION")
print("=" * 60)

print(f"\nFile: {os.path.basename(FILE_PATH)}")
print(f"Size: {os.path.getsize(FILE_PATH) / (1024**3):.2f} GB")


with Image.open(FILE_PATH) as img:

    print("\n--- BASIC INFORMATION ---")

    print("Format:", img.format)
    print("Mode:", img.mode)
    print("Size:", img.size)
    print("Number of TIFF pages:", getattr(img, "n_frames", "unknown"))

    print("\n--- TIFF INFORMATION ---")

    print("Number of bands:", len(img.getbands()))
    print("Bands:", img.getbands())

    print("\n--- TIFF TAGS ---")

    for tag_id, value in img.tag_v2.items():

        try:
            print(
                f"Tag {tag_id}: {value}"
            )
        except Exception:
            print(
                f"Tag {tag_id}: <unable to display>"
            )

    print("\n--- IMAGE ARRAY SHAPE ---")

    import numpy as np

    # Only load the current image plane
    arr = np.asarray(img)

    print("Array shape:", arr.shape)
    print("Array dtype:", arr.dtype)

print("\n" + "=" * 60)
print("DONE")
print("=" * 60)