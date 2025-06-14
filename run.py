from naming_analysis.controller import (
    setup_project_session,
    run_data_workflow,
    finalize_and_prompt
)

def main():
    # 🔹 1. Setup: Projekt, Konfiguration und Pfade
    book_name, config_data, data, paths, last_verse, mode_flags, naming_variants_dict = setup_project_session()

    # 🔹 2. Datenerhebung je nach Modus
    results = run_data_workflow(book_name, config_data, data, paths, last_verse, mode_flags, naming_variants_dict)

    # 🔹 3. Optional: Export und Analyse
    finalize_and_prompt(results, data, paths, book_name, config_data)


if __name__ == "__main__":
    main()