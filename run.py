from naming_analysis.controller import (
    setup_project_session,
    run_data_workflow,
    finalize_and_prompt
)

def main() -> None:
    """Run the naming-analysis workflow: setup → data processing → optional export/analysis."""
    # 1) Setup: project, configuration, and paths
    book_name, config_data, data, paths, last_verse, mode_flags, naming_variants_dict = setup_project_session()

    # 2) Data processing (depends on selected mode)
    results = run_data_workflow(data, paths, last_verse, mode_flags, naming_variants_dict)

    # 3) Optional: export and analysis prompt
    finalize_and_prompt(results, data, paths, book_name, config_data)

if __name__ == "__main__":
    main()