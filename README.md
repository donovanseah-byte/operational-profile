# Vessel Operating Profile App

## Install and run

Open a terminal in this folder, then run:

```bash
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

## A4 report

After uploading the Noon, Departure and Arrival workbooks and completing the inputs, open **A4 Professional Report**. Enter the optional preparer name and management comment, then select **Download one-page A4 PDF report**.

The report labels the FOC saving percentage as a user-entered assumption. It does not present the assumption as a measured retrofit result.

## Downloadable offline Windows app

The **Build offline Windows app** GitHub Actions workflow builds a one-folder
Windows application. In GitHub, open **Actions**, choose that workflow, select
**Run workflow**, and download `VesselOperatingProfile-Windows` from the
completed run's **Artifacts** section. Unzip the downloaded artifact, then unzip
`VesselOperatingProfile-Windows.zip`. Share the resulting
`VesselOperatingProfile` folder as a ZIP with colleagues. They extract the
*whole folder* and double-click `VesselOperatingProfile.exe`. A browser opens
to the app on their own PC; they can close the console window to stop it.

No Python installation or internet connection is needed to run this bundle.
The app listens only on `127.0.0.1` and processes uploaded workbooks on that
PC. Each new release requires another build and distribution. Some managed
Windows computers block unapproved executables; follow company software and
data-handling rules. Test on a second company PC with representative reports,
including the A4 PDF download, before wider use.
