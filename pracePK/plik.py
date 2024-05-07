# coding: utf-8
def read_and_process_files(file_list):
    data_frames = []
    for file in file_list:
        try:
            df = pd.read_excel(file, usecols="D, H:K", skiprows=1, names=['Imię i Nazwisko', 'Nazwa pracy', 'Termin oddania', 'Tag', 'Status'])
            data_frames.append(df)
        except Exception as e:
            print(f"Błąd otwarcia pliku '{file}': {e}")
    return data_frames

def summarize_work(data_frames):
    total_works = sum(len(df) for df in data_frames)
    reviewed_works = sum(df[df['Status'].isin(['Przesłano', 'Zwrócono'])].shape[0] for df in data_frames)
    return total_works, reviewed_works

results = []

for teacher, files in teachers_files.items():
    full_file_paths = [glob.glob(f"{f}")[0] for f in files if glob.glob(f"{f}")]
    if not full_file_paths:
        print(f"Nie znaleziono plików dla nauczyciela {teacher}")
    else:
        print(f"Pliki znalezione dla {teacher}: {full_file_paths}")
    data_frames = read_and_process_files(full_file_paths)
    total_works, reviewed_works = summarize_work(data_frames)
    percentage = (reviewed_works / total_works) * 100 if total_works else 0
    results.append({
        'Nauczyciel': teacher,
        'Ilość prac': total_works,
        'Ilość zwróconych': reviewed_works,
        'Procent': percentage
    })

results_df = pd.DataFrame(results)
print(results_df)
