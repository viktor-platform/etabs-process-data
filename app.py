import viktor as vkt
import pandas as pd
import plotly.graph_objects as go


def process_etabs_file(uploaded_file):
    # Read the file into a dataframe
    sheet_names = ["Joint Reactions", "Objects and Elements - Joints"]
    with uploaded_file.file.open_binary() as fr:
        dataframes = pd.read_excel(fr, sheet_name=sheet_names, skiprows=1)

    # Process the 'Joint Reactions' dataframe
    loads_df = dataframes["Joint Reactions"].dropna(subset=["Unique Name", "Output Case"]).copy()

    # Process the 'Objects and Elements - Joints' dataframe
    cords = dataframes["Objects and Elements - Joints"].dropna(
        subset=["Element Name", "Object Name", "Global X", "Global Y", "Global Z"]
    ).copy()
    cords = cords.rename(columns={"Object Name": "Unique Name"})

    # Get unique load case names as a list
    unique_output_cases = loads_df["Output Case"].unique().tolist()

    # Merge loads and cords dataframe
    merged_df = pd.merge(loads_df, cords, on="Unique Name", how="inner")

    return unique_output_cases, merged_df.reset_index(drop=True)


def get_load_combos(params, **kwargs):
    if params.xlsx_file:
        load_combos, _ = process_etabs_file(params.xlsx_file)  # Only use the first returned item, the load_combos
        return load_combos
    return ["First upload a .xlsx file"]


class Parametrization(vkt.Parametrization):
    intro = vkt.Text("""
# ETABS reaction heatmap
This app allows you to inspect results from an uploaded ETABS output file. Export your ETABS model results to an .xlsx file and upload it below. After uploading your Excel file, select the load combination you want to visualize.

Ensure the file includes the tables:
- **Joint Reactions**
- **Objects and Elements - Joints**
""")
    xlsx_file = vkt.FileField("Upload ETABS exported .xlsx")
    lb = vkt.LineBreak()
    selected_load_combo = vkt.OptionField("Select available load combos", options=get_load_combos)


class Controller(vkt.Controller):
    parametrization = Parametrization

    @vkt.PlotlyView("Heatmap")
    def plot_heat_map(params,**kwargs):
        if params.selected_load_combo:
            _, merged_df = process_etabs_file(params.xlsx_file)

            # Filter the dataframe based on the selected load combination
            filtered_df = merged_df[merged_df["Output Case"] == params.selected_load_combo]
            FZ_min, FZ_max = filtered_df["FZ"].min(), filtered_df["FZ"].max()

            # Create plotly scatter plot
            fig = go.Figure(
                data=go.Scatter(
                    x=filtered_df["Global X"],
                    y=filtered_df["Global Y"],
                    mode='markers+text',
                    marker=dict(
                        size=16,
                        color=filtered_df["FZ"],
                        colorscale=[
                            [0, "green"],
                            [0.5, "yellow"],
                            [1, "red"]
                        ],
                        colorbar=dict(title="FZ (kN)"),
                        cmin=FZ_min,
                        cmax=FZ_max
                    ),
                    text=[f"{fz:.1f}" for fz in filtered_df["FZ"]],
                    textposition="top right"
                )
            )

            # Style the plot
            fig.update_layout(
                title=f"Heatmap for Output Case: {params.selected_load_combo}",
                xaxis_title="X (m)",
                yaxis_title="Y (m)",
                plot_bgcolor='rgba(0,0,0,0)',
            )
            fig.update_xaxes(
                linecolor='LightGrey',
                tickvals=filtered_df["Global X"],
                ticktext=[f"{x / 1000:.3f}" for x in filtered_df["Global X"]],
            )
            fig.update_yaxes(
                linecolor='LightGrey',
                tickvals=filtered_df["Global Y"],
                ticktext=[f"{y / 1000:.3f}" for y in filtered_df["Global Y"]],
            )

            return vkt.PlotlyResult(fig.to_json())
        else:
            return vkt.PlotlyResult({})