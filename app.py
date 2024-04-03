import argparse
import fitz  # PyMuPDF
import xlsxwriter
from xlsxwriter import worksheet

OMIT_FIRST_LINES_COORDINATE = 85


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Extract text from PDF and write to an Excel file."
    )
    parser.add_argument("pdf_path", help="Path to the input PDF file")
    parser.add_argument("output_path", help="Path to the output Excel file")
    args = parser.parse_args()
    return args


def init_workbook(output_path):
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet()
    return workbook, worksheet


def write_text_boxes_to_excel(pdf_path, worksheet: worksheet.Worksheet):
    row = 1
    pdf_document = fitz.open(pdf_path)
    for page_number in range(len(pdf_document)):
        page = pdf_document.load_page(page_number)
        boxes = page.get_text_blocks()  # type:ignore
        text_boxes = [
            {
                "x0": box[0],
                "y0": box[1],
                "x1": box[2],
                "y1": box[3],
                "text": box[4].replace("\n", ""),
                "index": i,
            }
            for i, box in enumerate(boxes)
        ]
        grouped_text_boxes = group_dicts_by_range(text_boxes)
        for group in grouped_text_boxes:
            col = 0
            for element in sorted(group, key=lambda x: x["index"]):
                worksheet.write(
                    row, col, element["text"]
                )  
                col += 1
            row += 1
    pdf_document.close()


def group_dicts_by_range(dict_list):
    sorted_dicts = sorted(dict_list, key=lambda x: x["y0"])
    groups = []
    for d in sorted_dicts:
        added_to_existing_group = False
        for group in groups:
            if any(
                d["y0"] >= element["y0"] and d["y0"] <= element["y1"]
                for element in group
            ):
                if d["y1"] > OMIT_FIRST_LINES_COORDINATE:
                    group.append(d)
                    added_to_existing_group = True
                    break
        if not added_to_existing_group:
            groups.append([d])
    return sorted(groups, key=lambda x: x[0]["index"])


def main():
    args = parse_arguments()
    workbook, worksheet = init_workbook(args.output_path)
    write_text_boxes_to_excel(args.pdf_path, worksheet)
    workbook.close()


if __name__ == "__main__":
    main()
