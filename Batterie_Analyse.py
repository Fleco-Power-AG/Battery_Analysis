import os

import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
import pandas as pd

from oemof import solph


file_path = os.path.dirname(__file__)
filename = os.path.join(file_path, "pv_example_data.csv")
input_data = pd.read_csv(
    filename, index_col="timestep", parse_dates=["timestep"]
)

input_data.plot(drawstyle="steps")
plt.xlim(pd.Timestamp("2020-03-01 00:00"), pd.Timestamp("2020-03-07 00:00"))
plt.show()