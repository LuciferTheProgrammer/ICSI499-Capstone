import argparse
import os

import cv2


def crop_table_from_screenshot(
    image_path: str,
    output_path: str,
    padding: int = 10,
    white_threshold: int = 245,
    verbose: bool = True,
) -> bool:
    """
    Crop surrounding white space from an image and save the cleaned result.

    This function combines Otsu thresholding with a near-white mask so it can
    reliably detect both table borders and text-only content.
    """
    original_img = cv2.imread(image_path)
    if original_img is None:
        if verbose:
            print(f"Error: Could not load image at {image_path}")
        return False

    gray_img = cv2.cvtColor(original_img, cv2.COLOR_BGR2GRAY)

    _, otsu_mask = cv2.threshold(
        gray_img, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
    )
    _, near_white_mask = cv2.threshold(
        gray_img, white_threshold, 255, cv2.THRESH_BINARY_INV
    )
    content_mask = cv2.bitwise_or(otsu_mask, near_white_mask)

    non_zero = cv2.findNonZero(content_mask)
    if non_zero is None:
        if verbose:
            print("Error: No non-white content detected in the image.")
        return False

    x, y, w, h = cv2.boundingRect(non_zero)
    y_start = max(0, y - padding)
    y_end = min(original_img.shape[0], y + h + padding)
    x_start = max(0, x - padding)
    x_end = min(original_img.shape[1], x + w + padding)
    cropped_img = original_img[y_start:y_end, x_start:x_end]

    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    wrote_ok = cv2.imwrite(output_path, cropped_img)
    if not wrote_ok:
        if verbose:
            print(f"Error: Could not write image to {output_path}")
        return False

    if verbose:
        print(f"Success! Cropped image saved to: {output_path}")
    return True


def crop_table_in_place(
    image_path: str,
    padding: int = 10,
    white_threshold: int = 245,
    verbose: bool = True,
) -> bool:
    """Crop an image in place."""
    return crop_table_from_screenshot(
        image_path,
        image_path,
        padding=padding,
        white_threshold=white_threshold,
        verbose=verbose,
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Crop whitespace from screenshots.")
    parser.add_argument("input_image", help="Path to the input screenshot.")
    parser.add_argument(
        "-o",
        "--output",
        default="cropped_table.png",
        help="Path for the output image (default: cropped_table.png)",
    )
    parser.add_argument(
        "--padding",
        type=int,
        default=10,
        help="Padding (in pixels) to keep around detected content.",
    )
    parser.add_argument(
        "--white-threshold",
        type=int,
        default=245,
        help="Threshold used to classify near-white background pixels.",
    )

    args = parser.parse_args()

    if os.path.exists(args.input_image):
        crop_table_from_screenshot(
            args.input_image,
            args.output,
            padding=args.padding,
            white_threshold=args.white_threshold,
            verbose=True,
        )
    else:
        print(f"Error: Input file '{args.input_image}' does not exist.")