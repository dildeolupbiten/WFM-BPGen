# -*- coding: utf-8 -*-

import os
import json
import xlrd
import numpy as np
import pandas as pd
import tkinter as tk

from tkinter import ttk
from subprocess import Popen
from threading import Thread
from time import perf_counter
from PIL import Image, ImageTk
from xlsxwriter import Workbook
from tkcalendar import Calendar
from webbrowser import open_new
from urllib.error import URLError
from urllib.request import urlopen
from statistics import harmonic_mean, stdev
from tkinter.filedialog import askopenfilename
from datetime import datetime as dt, timedelta as td
from tkinter.messagebox import showinfo, showwarning, showerror
