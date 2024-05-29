from src.models.app_dv import App_DV


def main():
    import warnings
    warnings.filterwarnings("ignore")
    app = App_DV()
    app.crear_app()

if __name__ == "__main__":
    main()