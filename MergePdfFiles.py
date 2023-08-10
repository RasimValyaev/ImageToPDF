# pip install aspose-words
import os

import pandas as pd
import aspose.words as aw

image_path_source = r'C:\Users\Rasim\Desktop\ЕСП\Double\all'


def cycle_on_directory_files():
    df = pd.DataFrame(columns=["filename"])
    # цикл по папкам и файлам в папке pathName
    for root, dirs, files in os.walk(image_path_source):
        try:
            for i, name in enumerate(files):
                filename, file_extension = os.path.splitext(name.lower())
                if file_extension in ['.jpg', '.png', '.bmp'] and filename[:2] in ['рн', 'тт']:
                    # print(name)
                    df.loc[len(df)] = filename

        except Exception as e:
            print(str(e))
            continue

    df_set = df[df['filename'].str.contains("_") == False]
    df_set.reset_index(drop=True, inplace=True)
    print(df)
    print(df_set)
    return df, df_set


def merge_pdf(df, df_set):
    # fileNames = ["Input1.pdf", "Input2.pdf"]
    output = aw.Document()
    # Remove all content from the destination document before appending.
    output.remove_all_children()

    for unq_file_name in df_set.itertuples():
        df_filter = df[df['filename'].str.contains(unq_file_name.filename)].values.tolist()
        for file_name in df_filter:
            file_path = os.path.join(image_path_source, file_name[0] + '.jpg')
            input = aw.Document(file_path)
            # Append the source document to the end of the destination document.
            output.append_document(input, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)

        if df_filter != []:
            new_file_path = os.path.join(image_path_source, 'PDF', unq_file_name)
            print(new_file_path + ".pdf")
            output.save(new_file_path + ".pdf")


if __name__ == '__main__':
    df, df_set = cycle_on_directory_files()
    merge_pdf(df, df_set)
