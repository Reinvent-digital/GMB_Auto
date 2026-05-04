from pathlib import Path

from gmb_utils import (build_comparison_table, infer_month_label,
                       read_gmb_file, save_comparison_excel)


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Compare two GMB insights CSV exports side by side.")
    parser.add_argument("previous", help="Previous month export file path")
    parser.add_argument("current", help="Current month export file path")
    parser.add_argument("--third", help="Optional third month export file path")
    parser.add_argument("-o", "--output", help="Output CSV path", default="gmb_comparison.csv")
    parser.add_argument("--previous-label", help="Name for previous month (e.g. February)")
    parser.add_argument("--current-label", help="Name for current month (e.g. March)")
    parser.add_argument("--third-label", help="Name for third month (e.g. April)")
    args = parser.parse_args()

    datasets = []
    labels = []

    def add_file(path_str, custom_label):
        path = Path(path_str)
        label = custom_label or infer_month_label(path)
        df_val = read_gmb_file(path)
        datasets.append((df_val, label))
        labels.append(label)

    add_file(args.previous, args.previous_label)
    add_file(args.current, args.current_label)
    if args.third:
        add_file(args.third, args.third_label)

    result = build_comparison_table(datasets)
    if args.output.lower().endswith(('.xlsx', '.xls')):
        save_comparison_excel(result, args.output, labels)
    else:
        result.to_csv(args.output, index=False)

    print(f"Saved comparison to {args.output}")


if __name__ == "__main__":
    main()
