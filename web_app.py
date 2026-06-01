from app import DEFAULT_DASHBOARD_PORT, app


if __name__ == "__main__":
    print(f"Excel Arabic Dashboard  http://127.0.0.1:{DEFAULT_DASHBOARD_PORT}/")
    app.run(host="0.0.0.0", port=DEFAULT_DASHBOARD_PORT, debug=True)
