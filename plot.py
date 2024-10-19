# pyright: reportUnknownMemberType=false
# pyright: reportUnknownArgumentType=false
# pyright: reportPrivateImportUsage=false
# pyright: reportAttributeAccessIssue=false
# pyright: reportUnknownVariableType=false

# %%

import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

plt.style.use("default")

# %% Import data

s_sim = pd.read_csv("s_sim.csv")
s_meas = pd.read_csv("./20241011/s_2.csv", skiprows=2, sep=";").iloc[:, :-1]

# %% S-parameters plot

plt.style.use(["seaborn-v0_8-paper", "./publication.mplstyle"])

fig, ax = plt.subplots(figsize=(3.5, 2))

x_sim = np.linspace(22, 26, 61, endpoint=True)
x_meas = np.linspace(22, 26, 601, endpoint=True)

s11_sim = s_sim["dB(S(1,1)) []"].to_numpy()
s21_sim = s_sim["dB(S(2,1)) []"].to_numpy()
s22_sim = s_sim["dB(S(2,2)) []"].to_numpy()

s11_re_meas = s_meas["re:Trc1_S11"].to_numpy()
s11_im_meas = s_meas["im:Trc1_S11"].to_numpy()
s11_meas = s11_re_meas + 1j * s11_im_meas
s11_meas = 20 * np.log10(np.abs(s11_meas))
s21_re_meas = s_meas["re:Trc2_S21"].to_numpy()
s21_im_meas = s_meas["im:Trc2_S21"].to_numpy()
s21_meas = s21_re_meas + 1j * s21_im_meas
s21_meas = 20 * np.log10(np.abs(s21_meas))
s22_re_meas = s_meas["re:Trc3_S22"].to_numpy()
s22_im_meas = s_meas["im:Trc3_S22"].to_numpy()
s22_meas = s22_re_meas + 1j * s22_im_meas
s22_meas = 20 * np.log10(np.abs(s22_meas))


left = 24
bottom = -40
width = 0.25
height = 40
rect = plt.Rectangle(
    xy=(left, bottom),
    width=width,
    height=height,
    facecolor=mpl.cm.Paired(0),
    alpha=0.5,
)
ax.add_patch(rect)

ax.plot(
    x_sim,
    s11_sim,
    c="black",
    linestyle=(0, (4, 9)),
    label=r"sim. |$S_{11}$| & |$S_{22}$|",
)
ax.plot(
    x_sim,
    s11_sim,
    c=mpl.cm.Paired(5),
    linestyle=(6.5, (4, 9)),
)
# ax.plot(
#     x_sim,
#     s_sim["dB(S(2,1)) []"],
#     "-.",
#     c=mpl.cm.Paired(11),
#     label=r"sim. w/o decoupling",
# )
ax.plot(
    x_sim,
    s21_sim,
    linestyle="dotted",
    c=mpl.cm.Paired(9),
    label=r"sim. |$S_{21}$|",
)

ax.plot(
    x_meas,
    s11_meas,
    c="black",
    label=r"meas. |$S_{11}$|",
)
ax.plot(
    x_meas,
    s22_meas,
    c=mpl.cm.Paired(5),
    label=r"meas. |$S_{22}$|",
)
ax.plot(
    x_meas,
    s21_meas,
    c=mpl.cm.Paired(9),
    label=r"meas. |$S_{21}$|",
)

x_min = 22
x_max = 26
y_min = -40
y_max = 0

ax.set_xlabel("Frequency (GHz)")
ax.set_ylabel("Scattering Parameters (dB)")

ax.set_xlim(x_min, x_max)
ax.set_ylim(y_min, y_max)
ax.set_xticks(np.linspace(x_min, x_max, int((x_max - x_min) + 1), endpoint=True))
ax.set_yticks(np.linspace(y_min, y_max, int((y_max - y_min) / 10 + 1), endpoint=True))
ax.xaxis.set_minor_locator(mpl.ticker.MultipleLocator(1 / 2))
ax.yaxis.set_minor_locator(mpl.ticker.MultipleLocator(5))
ax.grid(which="minor", linestyle=":", alpha=0.5)

ax.legend(loc="lower left", prop={"math_fontfamily": "stix"})

fig.savefig("s-params.pdf", bbox_inches="tight")

# %% Import Data

p_sim = pd.read_csv("p_sim.csv")
p1_meas_e_co = pd.read_excel(
    "./20241015/board_nylon/cable_1_E_plane_Co.xlsx", header=[0, 1]
)
p1_meas_e_x = pd.read_excel(
    "./20241015/board_nylon/cable_1_E_plane_X.xlsx", header=[0, 1]
)
p1_meas_h_co = pd.read_excel(
    "./20241015/board_nylon/cable_1_H_plane_Co.xlsx", header=[0, 1]
)
p1_meas_h_x = pd.read_excel(
    "./20241015/board_nylon/cable_1_H_plane_X.xlsx", header=[0, 1]
)
p2_meas_e_co = pd.read_excel(
    "./20241015/board_nylon/cable_2_E_plane_Co.xlsx", header=[0, 1]
)
p2_meas_e_x = pd.read_excel(
    "./20241015/board_nylon/cable_2_E_plane_X.xlsx", header=[0, 1]
)
p2_meas_h_co = pd.read_excel(
    "./20241015/board_nylon/cable_2_H_plane_Co.xlsx", header=[0, 1]
)
p2_meas_h_x = pd.read_excel(
    "./20241015/board_nylon/cable_2_H_plane_X.xlsx", header=[0, 1]
)

# %% P1 E

plt.style.use(["seaborn-v0_8-paper", "./publication.mplstyle"])
fig, ax = plt.subplots(subplot_kw={"projection": "polar"}, figsize=(1.75, 1.75))

ax.set_theta_direction(-1)
ax.set_theta_zero_location("N")
ax.set_rlim(-40, 0)
ax.set_rticks(np.arange(-40, 1, 10))
ax.set_rlabel_position(45)
ax.tick_params(pad=2)
ax.set_thetagrids(
    np.arange(0, 360, 30),
    labels=[
        "θ=0°",
        "30°",
        "60°",
        "90°",
        "120°",
        "150°",
        "180°",
        "150°",
        "120°",
        "90°",
        "60°",
        "30°",
    ],
)

p1_e_theta = p_sim.iloc[:, 0] / 180 * np.pi
p1_e_co_sim = p_sim.iloc[:, 3]
p1_e_offset_sim = max(p1_e_co_sim)
p1_e_x_sim = p_sim.iloc[:, 1]

ax.plot(
    p1_e_theta,
    p1_e_x_sim - p1_e_offset_sim,
    "--",
    linewidth=1,
    c=mpl.cm.Paired(5),
    clip_on=False,
    zorder=2,
)
ax.plot(
    p1_e_theta,
    p1_e_co_sim - p1_e_offset_sim,
    "--",
    linewidth=1,
    c=mpl.cm.Paired(1),
    clip_on=False,
    zorder=3,
)

p1_e_co = p1_meas_e_co["Curve41"]["Amplitude (dB)"]
p1_e_offset = max(p1_e_co)
p1_e_x = p1_meas_e_x["Curve41"]["Amplitude (dB)"]

ax.plot(
    p1_e_theta,
    p1_e_x - p1_e_offset,
    c=mpl.cm.Paired(5),
    linewidth=1,
    clip_on=False,
    zorder=4,
)
ax.plot(
    p1_e_theta,
    p1_e_co - p1_e_offset,
    c=mpl.cm.Paired(1),
    linewidth=1,
    clip_on=False,
    zorder=5,
)

fig.savefig("p1_e.pdf", bbox_inches="tight")

# %% P1 H

plt.style.use(["seaborn-v0_8-paper", "./publication.mplstyle"])
fig, ax = plt.subplots(subplot_kw={"projection": "polar"}, figsize=(1.75, 1.75))

ax.set_theta_direction(-1)
ax.set_theta_zero_location("N")
ax.set_rlim(-40, 0)
ax.set_rticks(np.arange(-40, 1, 10))
ax.set_rlabel_position(45)
ax.tick_params(pad=2)
ax.set_thetagrids(
    np.arange(0, 360, 30),
    labels=[
        "θ=0°",
        "30°",
        "60°",
        "90°",
        "120°",
        "150°",
        "180°",
        "150°",
        "120°",
        "90°",
        "60°",
        "30°",
    ],
)
p1_h_theta = p_sim.iloc[:, 0] / 180 * np.pi

p1_h_co_sim = p_sim.iloc[:, 2]
p1_h_offset_sim = max(p1_h_co_sim)
p1_h_x_sim = p_sim.iloc[:, 4]
ax.plot(
    p1_h_theta,
    p1_h_x_sim - p1_h_offset_sim,
    "--",
    linewidth=1,
    c=mpl.cm.Paired(5),
    clip_on=False,
    zorder=2,
)
ax.plot(
    p1_h_theta,
    p1_h_co_sim - p1_h_offset_sim,
    "--",
    linewidth=1,
    c=mpl.cm.Paired(1),
    clip_on=False,
    zorder=3,
)

p1_h_co = p1_meas_h_co["Curve41"]["Amplitude (dB)"]
p1_h_offset = max(p1_h_co)
p1_h_x = p1_meas_h_x["Curve41"]["Amplitude (dB)"]

ax.plot(
    p1_h_theta,
    p1_h_x - p1_h_offset,
    c=mpl.cm.Paired(5),
    linewidth=1,
    clip_on=False,
    zorder=4,
)
ax.plot(
    p1_h_theta,
    p1_h_co - p1_h_offset,
    c=mpl.cm.Paired(1),
    linewidth=1,
    clip_on=False,
    zorder=5,
)

fig.savefig("p1_h.pdf", bbox_inches="tight")

# %% P2 E

plt.style.use(["seaborn-v0_8-paper", "./publication.mplstyle"])
fig, ax = plt.subplots(subplot_kw={"projection": "polar"}, figsize=(1.75, 1.75))

ax.set_theta_direction(-1)
ax.set_theta_zero_location("N")
ax.set_rlim(-40, 0)
ax.set_rticks(np.arange(-40, 1, 10))
ax.set_rlabel_position(45)
ax.tick_params(pad=2)
ax.set_thetagrids(
    np.arange(0, 360, 30),
    labels=[
        "θ=0°",
        "30°",
        "60°",
        "90°",
        "120°",
        "150°",
        "180°",
        "150°",
        "120°",
        "90°",
        "60°",
        "30°",
    ],
)

p2_e_theta = p_sim.iloc[:, 0] / 180 * np.pi
p2_e_co_sim = p_sim.iloc[:, 3]
p2_e_offset_sim = max(p1_e_co_sim)
p2_e_x_sim = p_sim.iloc[:, 1]

ax.plot(
    p2_e_theta,
    p2_e_x_sim - p2_e_offset_sim,
    "--",
    linewidth=1,
    c=mpl.cm.Paired(5),
    clip_on=False,
    zorder=2,
)
ax.plot(
    p2_e_theta,
    p2_e_co_sim - p2_e_offset_sim,
    "--",
    linewidth=1,
    c=mpl.cm.Paired(1),
    clip_on=False,
    zorder=3,
)

p2_e_co = p2_meas_e_co["Curve41"]["Amplitude (dB)"]
p2_e_offset = max(p2_e_co)
p2_e_x = p2_meas_e_x["Curve41"]["Amplitude (dB)"]

ax.plot(
    p2_e_theta,
    p2_e_x - p2_e_offset,
    c=mpl.cm.Paired(5),
    linewidth=1,
    clip_on=False,
    zorder=4,
)
ax.plot(
    p2_e_theta,
    p2_e_co - p2_e_offset,
    c=mpl.cm.Paired(1),
    linewidth=1,
    clip_on=False,
    zorder=5,
)

fig.savefig("p2_e.pdf", bbox_inches="tight")

# %%
# P2 H

plt.style.use(["seaborn-v0_8-paper", "./publication.mplstyle"])
fig, ax = plt.subplots(subplot_kw={"projection": "polar"}, figsize=(1.75, 1.75))

ax.set_theta_direction(-1)
ax.set_theta_zero_location("N")
ax.set_rlim(-40, 0)
ax.set_rticks(np.arange(-40, 1, 10))
ax.set_rlabel_position(45)
ax.tick_params(pad=2)
ax.set_thetagrids(
    np.arange(0, 360, 30),
    labels=[
        "θ=0°",
        "30°",
        "60°",
        "90°",
        "120°",
        "150°",
        "180°",
        "150°",
        "120°",
        "90°",
        "60°",
        "30°",
    ],
)
p2_h_theta = p_sim.iloc[:, 0] / 180 * np.pi

p2_h_co_sim = p_sim.iloc[:, 2]
p2_h_offset_sim = max(p1_h_co_sim)
p2_h_x_sim = p_sim.iloc[:, 4]
ax.plot(
    p2_h_theta,
    p2_h_x_sim - p2_h_offset_sim,
    "--",
    linewidth=1,
    c=mpl.cm.Paired(5),
    clip_on=False,
    zorder=2,
)
ax.plot(
    p2_h_theta,
    p2_h_co_sim - p2_h_offset_sim,
    "--",
    linewidth=1,
    c=mpl.cm.Paired(1),
    clip_on=False,
    zorder=3,
)

p2_h_co = p2_meas_h_co["Curve41"]["Amplitude (dB)"]
p2_h_offset = max(p2_h_co)
p2_h_x = p2_meas_h_x["Curve41"]["Amplitude (dB)"]

ax.plot(
    p2_h_theta,
    p2_h_x - p2_h_offset,
    c=mpl.cm.Paired(5),
    linewidth=1,
    clip_on=False,
    zorder=4,
)
ax.plot(
    p2_h_theta,
    p2_h_co - p2_h_offset,
    c=mpl.cm.Paired(1),
    linewidth=1,
    clip_on=False,
    zorder=5,
)

fig.savefig("p2_h.pdf", bbox_inches="tight")

# %%
gain_sim = pd.read_csv("gain_sim.csv")
gain_meas_p1_e = pd.read_excel(
    "./20241015/board_nylon/cable_1_E_gain.xlsx", header=[0, 1]
)
gain_meas_p1_h = pd.read_excel(
    "./20241015/board_nylon/cable_1_H_gain.xlsx", header=[0, 1]
)
gain_meas_p2_e = pd.read_excel(
    "./20241015/board_nylon/cable_2_E_gain.xlsx", header=[0, 1]
)
gain_meas_p2_h = pd.read_excel(
    "./20241015/board_nylon/cable_2_H_gain.xlsx", header=[0, 1]
)

# %% Gain

plt.style.use(["seaborn-v0_8-paper", "./publication.mplstyle"])

fig, ax = plt.subplots(figsize=(3.5, 2))

sim_x = np.linspace(22, 26, 41, endpoint=True)
meas_x = np.linspace(22, 28, 61, endpoint=True)

gain_p1 = gain_meas_p1_h["Curve1"]["Gain[dB]"]
gain_p2 = gain_meas_p2_h["Curve1"]["Gain[dB]"]

left = 24
bottom = 0
width = 0.25
height = 10
band_rect = plt.Rectangle((left, bottom), width, height, facecolor="skyblue", alpha=0.5)
ax.add_patch(band_rect)

ax.plot(
    sim_x,
    gain_sim.iloc[:, 1],
    "--",
    c="black",
    # linestyle=(0, (4, 9)),
    linewidth=1.0,
    label=r"sim. Port 1 & Port 2",
)
ax.plot(
    meas_x,
    gain_p1,
    c="black",
    linewidth=1.0,
    label=r"meas. Port 1",
)
ax.plot(
    meas_x,
    gain_p2,
    c=mpl.cm.Paired(5),
    linewidth=1.0,
    label=r"meas. Port 2",
)

ax.set_xlabel("Frequency (GHz)")
ax.set_ylabel("Realized Gain (dBi)")

ax.set_xlim(22, 26)
ax.set_ylim(0, 8)
ax.set_xticks(np.linspace(22, 26, 5, endpoint=True))
ax.xaxis.set_minor_locator(mpl.ticker.MultipleLocator(1 / 4))
ax.grid(which="minor", linestyle=":", alpha=0.5)

# font = fm.FontProperties(family="Times New Roman")
# ax.legend(loc="lower left", prop_family="Times New Roman")
ax.legend(loc="lower right", prop={"math_fontfamily": "stix"})

fig.savefig("gain.pdf", bbox_inches="tight")
