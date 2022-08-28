import numpy as np
import matplotlib.pyplot as plt

import japanize_matplotlib

def draw(title, X, labels):
    plt.figure(figsize=(10,10), dpi=100)
    plt.pie(X, labels=labels, autopct='%.f%%')
    plt.legend(labels, loc="right")
    plt.title(title)
    plt.savefig("figure.png", format="png", dpi=350)
    # plt.savefig('figure.png',bbox_inches='tight',pad_inches=0.05)
    plt.show()

