import tkinter as tk
import tkinter.messagebox
import math
from openpyxl import Workbook, load_workbook
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure


def reset():
    e_c_star.delete(0, "end")
    e_a.delete(0, "end")
    e_n.delete(0, "end")
    e_At.delete(0, "end")
    e_density.delete(0, "end")
    e_port_init.delete(0, "end")
    e_port_fin.delete(0, "end")
    e_length.delete(0, "end")
    e_Pa.delete(0, "end")
    e_epislon.delete(0, "end")
    e_gama.delete(0, "end")
    e_time_step.delete(0, "end")


Time = []
Port = [0]
Pc = []
r_dot = []
CF = []
F = []
Isp = []


def plot():
    for widget in Port_plot.winfo_children():
        widget.destroy()
    for widget in Pc_plot.winfo_children():
        widget.destroy()
    for widget in r_dot_plot.winfo_children():
        widget.destroy()
    for widget in CF_plot.winfo_children():
        widget.destroy()
    for widget in F_plot.winfo_children():
        widget.destroy()
    for widget in Isp_plot.winfo_children():
        widget.destroy()

    tk.Label(Port_plot, text="Time v.s. Port").pack()
    tk.Label(Pc_plot, text="Time v.s. Pc").pack()
    tk.Label(r_dot_plot, text="Time v.s. r-dot").pack()
    tk.Label(CF_plot, text="Time v.s. CF").pack()
    tk.Label(F_plot, text="Time v.s. F").pack()
    tk.Label(Isp_plot, text="Time v.s. Isp").pack()

    f1 = Figure(figsize=(3, 2), dpi=100)
    a = f1.add_subplot(111)
    a.plot(Time, Port)

    canvas1 = FigureCanvasTkAgg(f1, master=Port_plot)
    canvas1.draw()
    canvas1.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)

    f2 = Figure(figsize=(3, 2), dpi=100)
    b = f2.add_subplot(111)
    b.plot(Time, Pc)

    canvas2 = FigureCanvasTkAgg(f2, master=Pc_plot)
    canvas2.draw()
    canvas2.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)

    f3 = Figure(figsize=(3, 2), dpi=100)
    c = f3.add_subplot(111)
    c.plot(Time, r_dot)

    canvas3 = FigureCanvasTkAgg(f3, master=r_dot_plot)
    canvas3.draw()
    canvas3.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)

    f4 = Figure(figsize=(3, 2), dpi=100)
    d = f4.add_subplot(111)
    d.plot(Time, CF)

    canvas4 = FigureCanvasTkAgg(f4, master=CF_plot)
    canvas4.draw()
    canvas4.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)

    f5 = Figure(figsize=(3, 2), dpi=100)
    e = f5.add_subplot(111)
    e.plot(Time, F)

    canvas5 = FigureCanvasTkAgg(f5, master=F_plot)
    canvas5.draw()
    canvas5.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)

    f6 = Figure(figsize=(3, 2), dpi=100)
    f = f6.add_subplot(111)
    f.plot(Time, Isp)

    canvas6 = FigureCanvasTkAgg(f1, master=Isp_plot)
    canvas6.draw()
    canvas6.get_tk_widget().pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)

    toolbar1 = NavigationToolbar2Tk(canvas1, Port_plot)
    toolbar1.update()
    canvas1._tkcanvas.pack(
        side=tkinter.TOP,
        fill=tkinter.BOTH,
        expand=1,
    )

    toolbar2 = NavigationToolbar2Tk(canvas2, Pc_plot)
    toolbar2.update()
    canvas2._tkcanvas.pack(
        side=tkinter.TOP,
        fill=tkinter.BOTH,
        expand=1,
    )

    toolbar3 = NavigationToolbar2Tk(canvas3, r_dot_plot)
    toolbar3.update()
    canvas3._tkcanvas.pack(
        side=tkinter.TOP,
        fill=tkinter.BOTH,
        expand=1,
    )

    toolbar4 = NavigationToolbar2Tk(canvas4, CF_plot)
    toolbar4.update()
    canvas4._tkcanvas.pack(
        side=tkinter.TOP,
        fill=tkinter.BOTH,
        expand=1,
    )

    toolbar5 = NavigationToolbar2Tk(canvas5, F_plot)
    toolbar5.update()
    canvas5._tkcanvas.pack(
        side=tkinter.TOP,
        fill=tkinter.BOTH,
        expand=1,
    )

    toolbar6 = NavigationToolbar2Tk(canvas6, Isp_plot)
    toolbar6.update()
    canvas6._tkcanvas.pack(
        side=tkinter.TOP,
        fill=tkinter.BOTH,
        expand=1,
    )


def calculate():
    if (
        e_c_star.get() != ""
        and e_a.get() != ""
        and e_n.get() != ""
        and e_At.get() != ""
        and e_density.get() != ""
        and e_port_init.get() != ""
        and e_port_fin.get() != ""
        and e_length.get() != ""
        and e_Pa.get() != ""
        and e_epislon.get() != ""
        and e_gama.get() != ""
        and e_time_step.get() != ""
    ):

        def c_Port(t):
            return (
                Port[(int(t / time_step)) - 1]
                + 2 * r_dot[(int(t / time_step)) - 1] * time_step
            )

        def c_Pc(t):
            return (
                (math.pi * Port[int(t / time_step)] * length / At)
                * a
                * density
                * C_star
            ) ** (1 / (1 - n))

        def c_r_dot(t):
            return a * Pc[int(t / time_step)] ** n

        def pe_with_newton_method(Kn):
            def f(x):
                return (1 / x ** 2) * (
                    (2 / (gama + 1)) * (1 + (((gama - 1) / 2) * x ** 2))
                ) ** ((gama + 1) / (gama - 1)) - Kn ** 2

            def diff(x):
                return (-2 / x ** 3) * (
                    (2 / (gama + 1)) * (1 + (((gama - 1) / 2) * x ** 2))
                ) ** ((gama + 1) / (gama - 1)) + (1 / x ** 2) * (
                    (gama + 1) / (gama - 1)
                ) * (
                    (2 / (gama + 1)) * (1 + (((gama - 1) / 2) * x ** 2))
                ) ** (
                    ((gama + 1) / (gama - 1)) - 1
                ) * 2 * (
                    (gama - 1) / (gama + 1)
                ) * x

            difference = 1
            x0 = 10
            x = 0

            while difference > (1 / 10 ** 3):
                x = x0 - (f(x0) / diff(x0))
                difference = x0 - x
                x0 = x

            return Pa * (1 + ((gama - 1) / 2) * x0 ** 2) ** ((gama * -1) / (gama - 1))

        def c_CF(t):
            return (
                ((2 * gama ** 2) / (gama - 1))
                * ((2 / (gama + 1)) ** ((gama + 1) / (gama - 1)))
                * (1 - (Pe / Pc[int(t / time_step)]) ** ((gama - 1) / gama))
            ) ** (1 / 2) + (Pe - Pa) / Pc[int(t / time_step)] * epsilon

        def c_F(t):
            return CF[int(t / time_step)] * Pc[int(t / time_step)] * At

        def c_Isp(t):
            return C_star * CF[int(t / time_step)] / 9.81

        C_star = float(e_c_star.get())
        a = float(e_a.get())
        n = float(e_n.get())
        At = float(e_At.get())
        density = float(e_density.get())
        port_init = float(e_port_init.get())
        port_fin = float(e_port_fin.get())
        length = float(e_length.get())
        Pa = float(e_Pa.get())
        epsilon = float(e_epislon.get())
        gama = float(e_gama.get())
        time_step = float(e_time_step.get())
        Pe = pe_with_newton_method(epsilon)

        t = 0
        while Port[-1] <= port_fin:
            if t > 0:
                Port.append(c_Port(t))
            else:
                Port[0] = port_init
            Pc.append(c_Pc(t))
            r_dot.append(c_r_dot(t))
            CF.append(c_CF(t))
            F.append(c_F(t))
            Isp.append(c_Isp(t))
            Time.append(t)

            t += time_step

        plot()

        def output():
            wb = Workbook()
            ws = wb.active

            parameter = [
                "C*(m/sec)",
                "a(m/sec)",
                "n(Pa)",
                "At(m^2)",
                "Density(kg/m^3)",
                "Port_Initial_Diameter(m)",
                "Port_Final_Diameter(m)",
                "Fuel_Grain_Length(m)",
                "Atmospheric_Pressure(Pa)",
                "ε",
                "γ",
            ]
            ws.append(parameter)
            ws.append(
                [
                    C_star,
                    a,
                    n,
                    At,
                    density,
                    port_init,
                    port_fin,
                    length,
                    Pa,
                    epsilon,
                    gama,
                ]
            )
            ws.append([])

            title = [
                "Time(sec)",
                "Port(cm)",
                "Pc(Pa)",
                "r_dot(kg/sec)",
                "CF",
                "F(N)",
                "Isp(sec)",
            ]
            ws.append(title)

            for i in range(len(Time)):
                ws.append(
                    [
                        Time[i],
                        Port[i],
                        Pc[i],
                        r_dot[i],
                        CF[i],
                        F[i],
                        Isp[i],
                    ]
                )

            wb.save("data.xlsx")

        btn_output = tk.Button(window, text="Output data to data.xls", command=output)
        btn_output.pack(pady=10)

    else:
        tkinter.messagebox.showerror(title="ERROR", message="Don't leave blank!")


window = tk.Tk()
window.title("Solid Rocket Engine Simulator")
window.geometry("1200x750")

parameter_frame = tk.Frame(window)
parameter_frame.pack(padx=20, pady=20)

tk.Label(parameter_frame, text="Parameters:").grid(row=0, column=0, padx=10)

e_c_star = tk.Entry(parameter_frame, show=None, width=5)
e_a = tk.Entry(parameter_frame, show=None, width=5)
e_n = tk.Entry(parameter_frame, show=None, width=5)
e_At = tk.Entry(parameter_frame, show=None, width=5)
e_density = tk.Entry(parameter_frame, show=None, width=5)
e_port_init = tk.Entry(parameter_frame, show=None, width=5)
e_port_fin = tk.Entry(parameter_frame, show=None, width=5)
e_length = tk.Entry(parameter_frame, show=None, width=5)
e_Pa = tk.Entry(parameter_frame, show=None, width=5)
e_epislon = tk.Entry(parameter_frame, show=None, width=5)
e_gama = tk.Entry(parameter_frame, show=None, width=5)
e_time_step = tk.Entry(parameter_frame, show=None, width=5)

e_c_star.grid(row=1, column=1, padx=10, pady=10)
e_a.grid(row=1, column=3, padx=10, pady=10)
e_n.grid(row=1, column=5, padx=10, pady=10)
e_At.grid(row=1, column=7, padx=10, pady=10)
e_density.grid(row=1, column=9, padx=10, pady=10)
e_port_init.grid(row=1, column=11, padx=10, pady=10)
e_port_fin.grid(row=2, column=1, padx=10, pady=10)
e_length.grid(row=2, column=3, padx=10, pady=10)
e_Pa.grid(row=2, column=5, padx=10, pady=10)
e_epislon.grid(row=2, column=7, padx=10, pady=10)
e_gama.grid(row=2, column=9, padx=10, pady=10)
e_time_step.grid(row=2, column=11, padx=10, pady=10)

tk.Label(parameter_frame, text="C*(m/sec) =").grid(row=1, column=0)
tk.Label(parameter_frame, text="a(m/sec) =").grid(row=1, column=2)
tk.Label(parameter_frame, text="n(Pa) =").grid(row=1, column=4)
tk.Label(parameter_frame, text="At(m^2) =").grid(row=1, column=6)
tk.Label(parameter_frame, text="Density(kg/m^3) =").grid(row=1, column=8)
tk.Label(parameter_frame, text="Port_Initial_Diameter(m) =").grid(row=1, column=10)
tk.Label(parameter_frame, text="Port_Final_Diameter(m) =").grid(row=2, column=0)
tk.Label(parameter_frame, text="Fuel_Grain_Length(m) =").grid(row=2, column=2)
tk.Label(parameter_frame, text="Atmospheric_Pressure(Pa) =").grid(row=2, column=4)
tk.Label(parameter_frame, text="ε =").grid(row=2, column=6)
tk.Label(parameter_frame, text="γ =").grid(row=2, column=8)
tk.Label(parameter_frame, text="Time Step.(sec) =").grid(row=2, column=10)

btn_reset = tk.Button(parameter_frame, text="Reset", command=reset)
btn_reset.grid(row=5, column=5)
btn_confirm = tk.Button(parameter_frame, text="Confirm", command=calculate)
btn_confirm.grid(row=5, column=6)

plot_frame = tk.Frame(window)
plot_frame.pack()

Port_plot = tk.Frame(plot_frame)
Pc_plot = tk.Frame(plot_frame)
r_dot_plot = tk.Frame(plot_frame)
CF_plot = tk.Frame(plot_frame)
F_plot = tk.Frame(plot_frame)
Isp_plot = tk.Frame(plot_frame)

Port_plot.grid(row=1, column=0, padx=10)
Pc_plot.grid(row=1, column=1, padx=10)
r_dot_plot.grid(row=1, column=2, padx=10)
CF_plot.grid(row=3, column=0, padx=10)
F_plot.grid(row=3, column=1, padx=10)
Isp_plot.grid(row=3, column=2, padx=10)

tk.Label(window, text="Designer: 張詠翔").place(x=1100, y=730)

window.mainloop()
