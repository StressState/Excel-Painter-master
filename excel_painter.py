from PIL import Image
import argparse
import xlwings as xw

if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "input_file",
        help="input image."
    )
    parser.add_argument(
        "output_file",
        help="output xlsx filename."
    )
    parser.add_argument(
        "--mosaic",
        type=int,
        default=4,
        help="how many pixeles per cell."
    )
    parser.add_argument(
        "--max-width",
        type=int,
        default=256
    )

    args = parser.parse_args()
    original_image = Image.open(args.input_file)
    resize_image = original_image.resize(
        (args.max_width, int(float(original_image.size[1]) / original_image.size[0] * args.max_width)),
        Image.ANTIALIAS)

    wb = xw.Book(args.output_file)
    sht1 = wb.sheets['Sheet1']

    image_width = resize_image.size[0]
    image_height = resize_image.size[1]

    for row in range(1, image_height + 1):
        for col in range(1, image_width + 1):
            sht1.range((row,col)).column_width = 2.25
            color = resize_image.getpixel((col - 1, row - 1))
            print('position: ', row,',',col,'RGB: ', color)
            sht1.range((row,col)).color = color
