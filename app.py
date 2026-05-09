import streamlit as st
import fitz
import io
import zipfile
import re
import os
import base64

st.set_page_config(page_title="PDF智能填表系统 v14.4", layout="wide")

# ---- 字符宽度常量 ----
CHAR_W = 4.0
COMMA_W = 2.5
DOT_W = 2.0
FONTSIZE = 8.0

# ---- 内嵌SimSun子集字体（11.5KB base64，含逗号/小数点/星号/常用中文）----
COMMA_FONT_B64 = (
    "AAEAAAASAQAABAAgR1BPUwAZAAwAAC1oAAAAEEdTVULdBtr2AAAteAAAACBPUy8yUNL7"
    "AwAAGpwAAABgY21hcGlB0XkAABr8AAAArGN2dCAEugHNAAArnAAAALpmcGdtxWS09gAA"
    "G6gAAA3uZ2FzcABTADEAAC1UAAAAFGdseWb9SkRUAAABLAAAGGZoZWFk60DHUgAAGewA"
    "AAA2aGhlYQIBAOMAABp4AAAAJGhtdHgIywFCAAAaJAAAAFRsb2NhTgtUNAAAGbQAAAA4"
    "bWF4cAL8BNAAABmUAAAAIG5hbWUM8yhRAAAsWAAAANpwb3N0/+0ADAAALTQAAAAgcHJl"
    "cFFRD+cAACmYAAACBHZoZWEB4QDbAAAt0AAAACR2bXR4AWMAXgAALZgAAAA4AAEACQAb"
    "AHYAnQBJAIlAFyoiQBITAEwiQA0OAEwiJjEbDEAEJkcFuP/AtBITAEwFuP/Atg0OAEwF"
    "ASa4/8C0ExQATCa4/8C0DxAATCa4/8BAIAwATSYmS0pAASYbDAUfRAgfRC0tRB8IBBNA"
    "DQ4ATBM5AC/NKxc5Ly8vLxESFzkREgE5LysrK93NKysyEhc5EM0rKzIwMTcHFxYWFRQG"
    "IyImJycXFhQVFAYjIiY1NDc0NTcHBgYjIiY1NDY3NycmJjU0NjMyFhcXJzQ1JjU0NjMy"
    "FhUUBAcHNzY2MzIWFRQGZSIiDQQGAwMGBB8HAQYEAwcBCB8DBwMDBgYLIiIHCgYDAwcD"
    "HwcBBgQGBAEIHwQGAwQFBGkNDQUGAwgCAwMcKgMFAQYFBAcCAgIDKhwCBAQGBAYEDQ0D"
    "BQYFBQMDHCoDAgICBAcIAwIFAiocBAIGBAQFAAABAA7/4gAxACIAEAAlQBQPEA4QAEwJ"
    "IAowCgIKAA0GAAoJAwAvxDIBL9XFEMZdMjAxKzc0NjMyFhUUBgcnNjY1NCcmEgcHCAkQ"
    "DwQKCggIFgQIDQcOGAYFBg0HBQQEAAABABEAAQAvAB8ACwAOtAkDBgATAD/FAS/NMDE3"
    "IiY1NDYzMhYVFAYgBgkJBgYJCQEIBwcICQYHCAAAAgAMAAIAcwCxAAsAFwCHQBkLGBMV"
    "AEwLGAwNAEwHGBMVAEwHGAwNAEwFuP/otBMVAEwFuP/otAwNAEwBuP/otBMVAEwBuP/o"
    "QAsMDQBMCR0VAx0PFbj/wLMVAE0PuP/Asw4ATQ+4/8BAEgwATRUZDwwGSWwMDRIASWwSBQ"
    "A/Kz8rAS8QzisrKxDtEO0wMSsrKysrKysrNyIGFRQWMzI2NTQmByImNTQ2MzIWFRQGQA4"
    "TEw4PERAQFh4dFxYdHakrJCYqKiYkK6cuKicwLikpLwAAAQAdAAQAZgCyABYAQEATAR0M"
    "QBcATQxAExQATAxACwBNDLj/wEAWDABNDBIRagEHEhZzFgUHCGoHBWoHDQA/Kys/KysB"
    "LysrKyvtMDE3FRQWMzMVIzUzMjY1NTQmIyM1MzI2N0kKCQpJCQoJBQUSCQwPBLKYCAgG"
    "BggIeQQFBggIAAEADQAEAHAAsQApAL65ABX/4LQREwBMEbj/4LQREwBMBbj/2LQRFQBM"
    "KLj/6LMXAE0ouP/oswwATSK4//BAFxcATQIoFwBNAhgWAE0CGA8QAEwAHRYKuP/AthcA"
    "TQodEBy4//BACwwATRwkFhAWECMjuP/AswwATSG4/8BAIhUATSErIygDIyRLbBwgaiMc"
    "S2wYEwEjDRNzIw0TA0lsEwUAPys/KxI5KysrETkBLxDOKysROTkvLxEzKxDtKxDtMDEr"
    "KysrKysAKysrNzQmIyIGFRQXFhUUBiMiJjU0NjMyFhUUBgcGBgczMjY3MwcjNTY2NzY2"
    "XRAQDBEEAwUFBgceERgXCwwdHwM7CQwCBgZdBRcTERCFEhMNCgYEBAQFBQcIFRYYEgwY"
    "CxwjBg4OKw0JHBMRHgABAA8AAgBwALEAOADZuQA3/+izFwBNMrj/6LQREgBMLrj/6LMX"
    "AE0iuP/oQDwXAE0RGBcATS8QCwwATB8YFwBNHxgPEABMFEAXAE0UGBUWAEwUGA8QAEwE"
    "EAsMAEw2GRYdAB0dMycdLQy4/8BADxcATQwdBhkzLS0zGQMGBrj/wEAuDABNADoGNhoZ"
    "SWwaQAsMAEwaGgMwNiATAQMqMHMwIElsAxNJbAEDCTBzMAUDDQA/PysrKysREjkREjkv"
    "Kys5AS8QzisRFzkvLy8Q7SsQ7RDtEO0SOTAxKysrKysrKwArKysrKzcUBiMiJjU0NjMy"
    "FhUUBwYVFBYzMjY1NCYjNTI2NTQmIyIGFRQXFhUUBiMiJjU0NjMyFhUUBgcWFnAcGRIa"
    "CAQGBQICDwkSEhMaFBMPEAcQAgIEBQQHGhEUGA8QFhAyEx0UDQYJCQMFAwQDBggVFhIW"
    "BxMRDhUHCAgCAgMFBgYIDxMZDg8VBQgaAAIACQAEAHcAsQAUABgAZ7kAFf/wQBEXAE0P"
    "CBUATREPFwMLHRYNALj/wEAuEQBNDUANAE0AGg0YDxVNbA8OERZpDQ5JbBQNaQEGFBFJ"
    "DwF0DwUGB2oGBGoGDQA/Kys/KysrK4crADIBLxDUKysRM+0XOTAxKys3FBYzMxUjNTMy"
    "NjU1IzU3MxUzFSMnBzM1XAgHCUILCAdDSQobGxE5OhcHBgYGBgchBXRzBmBaWgAAAQAP"
    "AAIAcACvACcAoEAdCBALAE0bGBcATRsYDg8ATBcgFRYATBcgDg8ATAW4/+CzEABNAbj/"
    "4EAKEBIATBkdAyUgD7j/wEAMFwBNDx0JIiAiIAkJuP/AQCkMAE0DKQklHBYjJGkhI2og"
    "IQYfHAYWSWwBBgwhcwEGHABKIQF0IQQGDQA/PysrKxE5ERI5KysREjkBLxDOKxE5OS8v"
    "EO0rETMQ7TAxKysrKysrACs3MhYVFAYjIiY1NDYzMhYVFAcGFRQWMzI2NTQmIyIGByc3"
    "MwcjBzY2RRMYGRoSHAcGBgQDAxMLEBEPDwsQBwkFUgNIBAoTch0YGyAVDwYJCAMGAwQE"
    "CAgaGBQZCA0BWw9CDgYAAAIADAACAHMAsQALAC0A37kAKf/wswsATRq4/+i0CxIATBe4"
    "//BADwsATRMQFgBNExAOEABMEbj/8LQREgBMDbj/6LMXAE0NuP/wtBESAEwKuP/wQBEU"
    "FQBMCBAOEABMBBAOEABMArj/6LQUFQBMArj/8LQREwBMArj/8EAQDQBNIR0bGxUGHQ"
    "8AKx0VFbj/wLMOAE0VuP/AswwATQ+4/8BAIhUATQ8vFSsJAxgoSWwBEh4YcwESCQxKGA"
    "F0GAUSA0lsEg0APys/KysrERI5AS8QzisrKxDtMhDtEjkv7TAxKysrKysrKysrKysrKy"
    "s3FBYzMjY1NCYjIgY3MhYVFAYjIiY1NDYzMhYVFAYjIiY1NDc0NTQmIyIGFTY2HxUPDx"
    "ARDAsTHxgVHRMbHCIdDhEFBgUFAQYIERgGFUgeIRYeFBYNFyAUHh0pJy4xEAYGBgUFAg"
    "ICAgQFHzQNDQAAAQAQAAIAcQCvABcASLcWCBcATRAAB7j/wEAkFwBNBx0NQBcATQ1AEx"
    "QATA0NFQAZFRQVaREUagoNFhFLbBYEAD8rPysrAS8QzhE5Lysr7SsRMzAxKzcGBhUUFx"
    "QVFAYjIiY1NDY3IyIGByc3M3EVGQEFBgUHESUyCw0DBgdaqCJBGAQHCAoIBgcJEzhDDg"
    "4BKgADAAsAAgBxALEACwAjAC8BprUvMBcATSy4//BAFBYATSMQFBUATCAQFwBNFggX"
    "AE0JuP/wQCAQAE0IKBcATQgoFBUATAggFgBNCCAREwBMBzgUFgBMLLj/wLQVFgBMLLj/"
    "wLMTAE0suP/gsxcATSy4/+CzFABNLLj/4LQNEABMKLj/6LMXAE0ouP/oQA4QAE0mGBcA"
    "TSYYEABNIrj/+EAOEABNIhAMAE0UGBAATR24/+hACxESAEwZGBESAEwQuP/wQAoXAE"
    "0QGBESAEwOuP/oQBkREgBMChgQAE0IIAwATQgQCwBNBBgQAE0CuP/gsxcATQK4/+iz"
    "EABNArj/8EAaDg8ATCEVDBIkHR4GHQwqHRgAHRIeGB4YEhK4/8CzDgBNErj/wEAJDAB"
    "NEkALAE0MuP/AthUATQwxEi24//C0FhcATC24//CzFABNLbj/6EAZEwBNCRgXAE0tIR"
    "UJAxsnSWwbBQ8DSWwPDQA/Kz8rETk5OTkrKysrAS8QzisrKysROTkvLxDtEO0Q7RDtER"
    "I5OTAxKysrKysrKysrKysrKysrKysrKysrKysrACsrKysrKysrKysrNxQWMzI2NTQmJwY"
    "GFxQGIyImNTQ2NyYmNTQ2MzIWFRQGBxYWJzQmIyIGFRQWFzY2GRUQExEYGAwNWBwXFR4"
    "REA4PGhcVGg8QERAQEBETEBcVDAwtEBQUDw0YCggYDBMaGhERGAkHFQ0RGBgRDBUICh"
    "hKDhUUCwsVCAYSAAACAAsAAgByALEACwAsAPi5ABb/6EApEABNDhAXAE0pEBYATSgY"
    "FgBNKBgOEgBMKBgLAE0aGAsSAEwXGAsATRO4//CzFgBNE7j/+EAVDhAATA0YFwBNDR"
    "gREgBMDQgLAE0KuP/osxYATQq4/+izEABNCrj/+EAKDg8ATAQQEwBNArj/6LMQAE0C"
    "uP/4QA8ODwBMBCodFSEdGwAdDxW4/8CzFQBND7j/wLMOAE0PuP/AQCcMAE0PQAsATRU"
    "uDyoJAxgnSWwBGB4ScxIJSWwBGAwDShIBdBIFGA0APz8rKysrERI5AS8QzisrKysQ7dT"
    "tEO0yMDErKysrKysrKysrKysrKysrKwArKzcUFjMyNjU0JiMiBhciJjU0NjMyFhUUBi"
    "MiJjU0NjMyFhUUFhcWFjMyNjcGBhwRDQ8WEBIREBsUGBoYGRwhHQ0TBQYGBAIBAQUDE"
    "hYCBRR6FhccDBgdG0sZHBggKCYqNwwKBAgJBgQFAQEBKC8MEAAAAQAM/+oA9QDPABUA"
    "DLMJEQARAC8vETMwMTcXBx4CFxUGBy4CJw4CByc+Am8eCgYdMR4UCBsjFgUEFDAqAi"
    "orD88NCTVKMwUFAQ0TMzozMEE1EgMaPV0AAAQACv/mAPUAzwASAC0AOQBCAB9ADSYaDD"
    "86LiITLCwDLgMALy8SOS8zMxEzMjMyMjAxNxQXBzY1NQYHJzY2NxcHBgcXBxczNxcHBg"
    "YHJic1FjY2NyMOAgcnPgI3Iwc3FwceAhcVBgcmJicXBwYGByc2NkEBEAESFQIWIQkWCR"
    "IKDgglVAkNCAIFFQQbGwwEAywDDBwfARoYCQEGDDgRCAYXHA0SAxkdKRUJEB8UAxYfFB"
    "kPBhcSaBkTBBlCIA0FIhMHBxUJDQZFHgkNCAQEAgtOICgeDgQOIisVAncHBhwkFAQDAw"
    "gSPhYNBSUqEAMVQQAAAgAP/+YA8QDSAC4ANwAfQA0PNC8YJSsiKxgYHCgcAC8vEjkvM"
    "zMRMxEzMjIwMTcHFhYVFAYjIicmJw4CBzQnNxYWNjY3Iw4CByc+AjcjIgcnMzY3FwcHM"
    "zcXBzMGBiY1NDc2wgIuAwYBBAUIGgQKEhIhARsQDQoEPwQRJycCISMPAyUNCwlGAgEX"
    "BwQ9ChGSBQQVDQkPih4cDgQICQ0SGjcwEwUMDwUIAQs6TCU4NRcDGTY0IwMJHh8MBisL"
    "DiIsDwQDBAcMAAAEAA//6QD1ANAACAA1AEwAUQAlQBBBTTZLSwMPNSshIRcXRQNFAC"
    "8vEjkvMxEzMzMROS8zMzIwMTc2NjcXBgcGBxc2NxcHBgczNxcGByc3IxYGJjU0NzY3Mx"
    "czJiYnNxYWFRQHMyYmJzcWFhUUBwczNxcHBgcWNxUGByYnBgcnNjcmJyMHNxYXNjcnLW"
    "ASEC4lJjZ6DQoXCxELLQkUEhQDDakCDA8GCgQEARcBBgsDDw8ILgIGCAMNDwlMcgoSDB"
    "IZK0AUAjUrMD8BPycXDA4JHBAYGQ60Aw8KFAQDBAE0FyAPBBUPCxQCGwIeFQwDAwMHDB"
    "gNBhUPAggQBQcIDhIOAggSBggIJQoRBRwXHAQEBQsGHyAIBA4eGR8CAhwVFhsAAAMAEP"
    "/oAPAAzgAOABIAMAArQBMtICANLCgmJg0QBQUNCw8PFw0XAC8vEjkvMxE5LzMROS8zMx"
    "E5LzMwMTcUFwc1IxUHNjQnFzM3FwcVMzUXDgIHJic3FjMyNjcjByc2NjcjByczNxcjBz"
    "M3F8ABD2QPAQEPYwkOemQUCQ4OEAIdARsNBQ8IXAkMBwkEOQkIwQ0ShQ5ZCw+7LQ0GCg"
    "kFFC0UCAoMBDMzhzQQBwQPDAQIBjcIDgQSDgIIDRMkCw8AAAIAC//mAPUA0wADAC4AM"
    "UAWASESHhIWFgQlDQANEREEDAgoKBoEGgAvLxI5LzMzETkvMzMRMxE5LzMzETMzMDE3"
    "FTM1JxcHBgczNxcjFTM3FyMVMzcXIxUUFwc2NTUjIgcnMzQnFzM1IwYGByc2Nkc6OBYJ"
    "CAiCDxVcKQ4US0EPFWUBEQFVDQsJLQEQOj4KFhMDEiBrNDRoDgQMDA8VMg4UNA8VJhEN"
    "Bx8OHgMJLhUJMg8ZDwITMwADADb/7QDSAMgAAwAHABcAJ0ARFA4LBRYWCwEEBAsMAAAI"
    "CwgALy8SOS8zETkvMxE5LzMRMzIwMTcVMzUHFTM1BzY0JxczNxcHFRQXBzUjFUV3d3eG"
    "AQEPdQcRCAEPd7hMTFJXV3k7Yz0KCg0GiR0YCBoTAAMAEP/mAMwAygADAAgAIQAjQA8f"
    "Bg4OGgEEBBoYAAARGhEALy8SOS8zETkvMxE5LzMyMDE3FTM1BxQHMzUHFjY1NSMGBgcn"
    "PgMnFzM3FwcVFAcmJ1RhYQFiLhsTYgMlGQIUFgsBAQ9fCRAJGQEiuTIyOBcaMX0EAQw9"
    "KC4OAxAgK0VACgsOB7ISCA4IAAAJAA3/5wDvAM8AAwAHAAsAJwAvADoAPgBDAFwAXUAs"
    "WhQONjAsKCAJIiIgIA5BSUkOBQgIDjw/Pw4BBAQOJxEbJREAAA5TOztMDkwALy8SOS8z"
    "ETkvMzMzETMROS8zETkvMxE5LzMROS8zETMRMzIzMhEzMjAxNxUzNQcVMzUHFTM1JzQ"
    "nFwcVMzQnFwcVMzcXIxUzNxcjByczNSMHJxcXBwYHJzY2NxYWFRQGIyInJic3FTM1BxQ"
    "HMzUHFhY1NSMGBgcnPgI0JxczNxcHFRYHNCc/KSkpKSk3ARcIKQEXCAMLER8BCxFtDggh"
    "CQ4ILBIJGx0CDxkqGAgGAgMDAwUPTDExATIkExEyAiEZAhEWCgEPLwcQBwEXHJ8aGiAe"
    "HiQfH0oZEQ0GFxIYDAYYCxFjCxECCGMCCHMQAyASAw4gEA0KBAUKCA0TijIyOBwXM3wD"
    "AQs8IywPAw4hLGYaCQoMB64UCA0IAAYADf/0APcAxAAcAEIARgBKAE4AUgBfQC0jIDE4"
    "OD4sOywwMCBQTCtMQUEgBxIEEhYWIEhET0RLSyARDQoKIEchQ0M1IDUALy8SOS8zMxE5"
    "LzMzETkvMzMRMxE5LzMzETMROS8zMxEzETkvMzMRMzIRMxEzMDE3Njc1IyIHJzM1IyIH"
    "JzM3FyMVMzcXIxU3FwYHBzc2NCcXMzcXBxUUFwc1IxUzNxcjFTM3FyMiByczNSMiBycz"
    "NSMVNRUzNTMVMzUHFTM1MxUzNQ0eCwUNCwkmBw0LCT8MEScKDBEnKQEmJghXAQENUwcP"
    "CAENJB0METoqDRKIDQsJVBYNCwk3JSUMJFUlDCQXCARDAwk9AwkMEj0MEj8OBBETCEMg"
    "PCAJCQwGLyASBgwrDBIpDRMDCSkDCSsKaCgoKCguKioqKgAABAAM//UA9QDNABkAKwBGA"
    "E0AJ0ARR0opHyYmHhopKTwWExMjPCMALy8SOS8zETkvMzMyETMRMzIwMTcGBx4CBiMiJ"
    "yYnBgYHJzY2NyMiByczNxcHMzcXIxUzNxcjIgcnMzUjIgcHNxcGBgcnNjY3BgYHJzY2N"
    "xcHBgc3NjcXBwYHJzY3FwYGxw4PIRwDBQIEBwonCiMfASEqEi0NCwlPCg9jWAwROi8M"
    "EYANCwlHDA0LU0EBEycNDAodDw8YDQkMIQUVCRwUMg0HFAsgMgoxLgESObMTEw0QEAcK"
    "DhgLGg0EEiogAwkKD3AMEkQMEgMJRAMCCwQFDQkTAyEXAwYIEwM8FQ0GLhMDFBIPBC5p"
    "EgkLBAYUAAADAAr/9QD1AMUACQAcADkANUAYEAYVAR8nNC4uIiYmHxgVFR8hHTc3Kx8r"
    "AC8vEjkvMzMROS8zETkvMzIRMzMRMxEzMjAxNzcWFgYGIyInJhc3FwYHBgcnNic1IyIH"
    "JzM3Fwc3MzcXIxUzNxcjFTM3FyMiByczNTUXBxUzNSMiByECEg4DBwEEAQMLIQMQCAkG"
    "DgUBCQcLCSIHEQkcdhATRh4PE0ArDxGSBwsJJRYJIyoIC78CCg4NBgkLlyACFgsLDg8F"
    "DVgDCQoMB0kPFVAOFFwOFAMJaBwMB3GyAwAEABH/5gDtANMAAwAHAA0ANQAxQBYNJyck"
    "JC4FDAwuAQQELjErKwAAFi4WAC8vEjkvMxEzETkvMxE5LzMROS8zETMwMTcVMzUHFTM1"
    "BzY3NSMVNzY3FwYHFRQHJic1FjYnNSMGByc2NyMiByczNCcXMzY3FwYHMzcXB1NhYWESC"
    "wdhbxUGEAkiFwEdGQ8BBUlUAVI6bAgLCTABDxMNBBMIFkYIEAioHR0jHBxACgkLHh4aDR"
    "EBKFATBw8GBAIBCDlCEQMbNQMJUx8JGA0MARgLDQYAAAAAAQAAABsBgQAxAAAAAAACAB"
    "AALwCIAAACFwMeAAAAAAAAAAAAqgDaAPcBYAGhAjwC9ANMA9UEhATNBecGowbPB0gHrA"
    "g9CJ0I+QkzCXoKLwrPC1oLyQwzAAEAAAAFJmZ5ZOj+Xw889QALAQAAAAAAt5hCgAAAAA"
    "DThx+z//7/3AEAANwAAAAMAAIAAAAAAAABAAAAAIAACQCAAA4AgAARAIAADACAAB0AgA"
    "ANAIAADwCAAAkAgAAPAIAADACAABAAgAALAIAACwEAAAwACgAPAA8AEAALADYAEAANA"
    "A0ADAAKABEAAQAAANz/3AAkAQD//v/3AQAAAQAAAAAAAAAAAAAAAAAAAA8AAwCAAZAABQ"
    "AIAIAAgAAAABAAgACAAAAAgAAMAEEAAAIBBgADAQEBAQEAAAABCAAAAAAAAAAAAAAAW"
    "llFQwBAACqOqwDc/9wAJADcACQAAAABAAAAAAB0AK8AAAAgAAEAAAABAAMAAQAAAAwAB"
    "ACgAAAAJAAgAAQABAAqACwALgA5TrpO/VKeU9dT9150ZeVnCGcfdAZ+z4vBjqv//wAAAC"
    "oALAAuADBOuk79Up5T11P3XnRl5WcIZx90Bn7Pi8GOq////9f/1v/V/9SxVLESrXKsOqw"
    "boZ+aL5kNmPeMEYFJdFhxbwABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAABARoeGhYSDgoGAf359fHt6eXh3dnV0c3JxcG9ubWxramloZ2ZlZGNiYWBfXl"
    "1cW1pZWFdWVVRTUVBPTk1MS0pJSEdGKB8QCgksAbELCkMjQ2UKLSwAsQoLQyNDCy0sAb"
    "AGQ7AHQ2UKLSywTysgsEBRWCFLUlhFRBshIVkbIyGwQLAEJUWwBCVFYWSKY1JYRUQbIS"
    "FZWS0sALAHQ7AGQwstLEtTI0tRWlggRYpgRBshIVktLEtUWCBFimBEGyEhWS0sS1MjS1"
    "FaWDgbISFZLSxLVFg4GyEhWS0ssAJDVFiwRisbISEhIVktLLACQ1RYsEcrGyEhIVktLL"
    "ACQ1RYsEgrGyEhISFZLSywAkNUWLBJKxshISFZLSwjILAAUIqKZLEAAyVUWLBAG7EBAy"
    "VUWLAFQ4tZsE8rWSOwYisjISNYZVktLLEIAAwhVGBDLSyxDAAMIVRgQy0sASBHsAJDIL"
    "gQAGK4EABjVyO4AQBiuBAAY1daWLAgYGZZSC0ssQACJbACJbACJVO4ADUjeLACJbACJWC"
    "wIGMgILAGJSNiUFiKIbABYCMbICCwBiUjYlJYIyGwAWEbiiEjISBZWbj/wRxgsCBjIyE"
    "tLLECAEKxIwGIUbFAAYhTWli5EAAAIIhUWLICAQJDYEJZsSQBiFFYuSAAAECIVFiyAgI"
    "CQ2BCsSQBiFRYsgIgAkNgQgBLAUtSWLICCAJDYEJZG7lAAACAiFRYsgIEAkNgQlm5QAAA"
    "gGO4AQCIVFiyAggCQ2BCWblAAAEAY7gCAIhUWLICEAJDYEJZsSYBiFFYuUAAAgBjuAQA"
    "iFRYsgJAAkNgQlm5QAAEAGO4CACIVFiyAoACQ2BCWVlZWVlZsQACQ1RYQAoFQAhACUAM"
    "Ag0CG7EBAkNUWLIFQAi6AQAACQEAswwBDQEbsYACQ1JYsgVACLgBgLEJQBuyBUAIugGA"
    "AAkBQFm5QAAAgIhVuUAAAgBjuAQAiFVaWLMMAA0BG7MMAA0BWVlZQkJCQkItLEWxAk4r"
    "I7BPKyCwQFFYIUtRWLACJUWxAU4rYFkbIyGwAyVFIGSKY7BAU1ixAk4rYBshWVlELSwg"
    "sABQIFgjZRsjWbEUFIpwRbBPKyOxYQYmYCuKWLAFQ4tZI1hlWSMQOi0ssAMlSWMjRmCwT"
    "ysjsAQlsAQlSbADJWNWIGCwYmArsAMlIBBGikZgsCBjYTotLLAAFrECAyWxAQQlAT4AP"
    "rEBAgYMsAojZUKwCyNCsQIDJbEBBCUBPwA/sQECBgywBiNlQrAHI0KwARaxAAJDVFhFI"
    "0UgGGmKYyNiICCwQFBYZxtmWWGwIGOwQCNhsAQjQhuxBABCISFZGAEtLCBFsQBOK0Qt"
    "LEtRsUBPK1BbWCBFsQFOKyCKikQgsUAEJmFjYbEBTitEIRsjIYpFsQFOKyCKI0REWS0s"
    "S1GxQE8rUFtYRSCKsEBhY2AbIyFFWbEBTitELSxLUrEBAkNTWlgjECABPAA8GyEhWS0s"
    "I7ACJbACJVNYILAEJVg8GzlZsAFguP/pHFkhISEtLLACJUewAiVHVIogIBARsAFgiiAS"
    "sAFhsIUrLSywBCVHsAIlR1QjIBKwAWEjILAGJiAgEBGwAWCwBiawhSuKirCFKy0AAED/"
    "XDMaH1szQB9aM/8fWTL/H1gxgB9XMUAfVjD/H1UwKx9UL/8fUy0gH1IuQB9RLv8fUCz/"
    "H08sKx9OKisfTSr/H0wp/x9LKBAfSigrH0ko/x9IJ0AfRyf/H0Ym/x9FJf8fRCSAH0Mk"
    "gB9CIxofQSOAH0AjgB8/IkAfPiL/Hz0iQB88If8fOyD/Hzof/x85Hv8fOB0WHzcdKx82H"
    "f8fNR1AHzQc/x8uLYAfLSuAHywrIB8lGf8fJAgbGVwjCBoZXCIZ/x8hFv8fIAwYFlwfF"
    "w0fHhf/Hx0W/x8cFg0fGxsZAFsYGBYAWxobGQBbFxgWAFsVGTgWOFoPFQH/FQETTRJVQE"
    "gR/xBVElkQWQ1NDFUFTQRVDFkEWQ+ADlULTQpVB00GVQEQAFUOWQpZBlkAWQlNCFUDTQ"
    "JVCFkCWSACUAKAArAC4AIFA0BABQG5AZAAVCtLuAf/UkuwCFBbsAGIsCVTsAGIsEBRWr"
    "AGiLAAVVpbWLEBAY5ZhY2NAB1CS7CQU1iyAwAAHUJZsQICQ1FYsQQDjllzACsAKysrAC"
    "sAKwArACsrKysrACsAKysrACsAKysrAXN0ASsBKwErASsBKysrACsrASsrASsAKwErAC"
    "sBKysrKysrKysAKysrKwErKysAKysrKysrASsrKwArKysrKysBKysrKysrACsrKysrKy"
    "srGAC3//gArwACAK8AAgB0AAIAAAACAAAAAgAAAAL/3///ALEAAAAAAAIAAAAPABAABg"
    "AGAA4ABgAGAAYAEAAGABAACAAQAA4ABwAHAA4ACQARAAcAGQAHAA0ACQAOAA0AFwAJAB"
    "UAAgAJAAYADgAQABMAFgAGABAACAAQAAwADgAQAAYACAAMAAYACAAOAAkAEQATAAcACg"
    "APABkABwAKAAYACQANAA8AEgAXAAYACQATABYAAgAJAAsADgAAAAAABwBaAAMAAQQJAA"
    "AATgAAAAMAAQQJAAEADABOAAMAAQQJAAIADgBaAAMAAQQJAAMADABOAAMAAQQJAAQA"
    "DABOAAMAAQQJAAUAGABoAAMAAQQJAAYADABOAKkAIABDAG8AcAB5AHIAaQBnAGgAdAA"
    "gAFoASABPAE4ARwBZAEkAIABFAGwAZQBjAHQAcgBvAG4AaQBjACAAQwBvAC4AIAAyAD"
    "AAMAAxAFMAaQBtAFMAdQBuAFIAZQBnAHUAbABhAHIAVgBlAHIAcwBpAG8AbgAgADUALg"
    "AxADUAAAADAAAAAAAA/+oADAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEABAAHAAoAEQAF"
    "ADsAD///AA8AAQAAAAoADAAOAAAAAAAAAAEAAAAKABwAHgACaGFuaQAObGF0bgAOAAAA"
    "AAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANAA0ACgAMAA4ACQAU"
    "ABIADQAYAA8AFwAJ"
)


def extract_font_data(doc):
    """从PDF提取SimSun字体数据"""
    for xref in range(1, doc.xref_length()):
        try:
            obj = doc.xref_object(xref)
            if not obj:
                continue
            if "/FontFile2" in obj:
                match = re.search(r'/FontFile2\s+(\d+)', obj)
                if match:
                    ff2_xref = int(match.group(1))
                    data = doc.xref_stream(ff2_xref)
                    if data and len(data) > 1000:
                        return data
        except:
            continue
    for xref in range(2, min(20, doc.xref_length())):
        try:
            data = doc.xref_stream(xref)
            if data and len(data) > 10000:
                return data
        except:
            continue
    return None


def make_comma_font():
    """从内嵌base64生成包含逗号/小数点/星号的SimSun子集字体"""
    subset_path = '/tmp/simsun_subset_comma_v14.ttf'
    if os.path.exists(subset_path) and os.path.getsize(subset_path) > 1000:
        return subset_path
    try:
        font_data = base64.b64decode(COMMA_FONT_B64)
        with open(subset_path, 'wb') as f:
            f.write(font_data)
        st.info(f"✅ 内嵌字体释放成功: {len(font_data)} bytes")
        return subset_path
    except Exception as e:
        st.error(f"❌ 内嵌字体释放失败: {e}")
        return None


def fmt_decimal(value, field_key):
    """格式化数字"""
    if not value or not str(value).strip():
        return value
    if field_key.startswith("eq"):
        return str(value).strip()
    if field_key in ["agent_name", "agent_id", "receiver", "receive_date"]:
        return str(value).strip()
    text = str(value).strip().replace(",", "")
    if "." in text or "%" in text:
        return text
    try:
        return "{:,.2f}".format(int(text))
    except:
        return text


FIELD_CFG = {
    # --- 第1行: 从业人数 ---
    "eq1s": (0, 100.6, 152.3, 164.8, 183.6, 0.5),
    "eq1e": (0, 152.6, 204.0, 164.8, 183.6, 0.5),
    "eq2s": (0, 204.3, 255.7, 164.8, 183.6, 0.5),
    "eq2e": (0, 256.0, 307.4, 164.8, 183.6, 0.5),
    "eq3e": (0, 514.5, 565.9, 164.8, 183.6, 0.5),
    # --- 第2行: 资产总额 ---
    "aq1s": (0, 100.6, 152.3, 183.6, 202.3, 0.5),
    "aq1e": (0, 152.6, 204.0, 183.6, 202.3, 0.5),
    "aq2s": (0, 204.3, 255.7, 183.6, 202.3, 0.5),
    "aq2e": (0, 256.0, 307.4, 183.6, 202.3, 0.5),
    "aq3e": (0, 514.5, 565.9, 183.6, 202.3, 0.5),
    # --- 预缴税款计算 ---
    "L1": (0, 514.5, 565.9, 297.7, 305.7, 0.2),
    "L2": (0, 514.5, 565.9, 316.5, 324.5, 0.2),
    "L3": (0, 514.5, 565.9, 335.2, 343.2, 0.2),
    "L4": (0, 514.5, 565.9, 354.0, 362.0, 0.2),
    "L5": (0, 514.5, 565.9, 372.7, 380.7, 0.2),
    "L6": (0, 514.5, 565.9, 391.5, 399.5, 0.2),
    "L7": (0, 514.5, 565.9, 410.2, 418.2, 0.2),
    "L8": (0, 514.5, 565.9, 429.0, 437.0, 0.2),
    "L9": (0, 514.5, 565.9, 447.7, 455.7, 0.2),
    "L10": (0, 514.5, 565.9, 466.5, 474.5, 0.2),
    "L11": (0, 514.5, 565.9, 485.2, 493.2, 0.2),
    "L12": (0, 514.5, 565.9, 504.0, 512.0, 0.2),
    "L13": (0, 514.5, 565.9, 522.7, 530.7, 0.2),
    "L13_1": (0, 514.5, 565.9, 541.5, 549.5, 0.2),
    "L14": (0, 514.5, 565.9, 560.2, 568.2, 0.2),
    "L15": (0, 514.5, 565.9, 579.0, 587.0, 0.2),
    "L16": (0, 514.5, 565.9, 597.7, 605.7, 0.2),
    "L17": (0, 514.5, 565.9, 631.5, 639.5, 0.2),
    "L18": (0, 514.5, 565.9, 650.2, 658.2, 0.2),
    "L19": (0, 514.5, 565.9, 669.0, 677.0, 0.2),
    "L20": (0, 514.5, 565.9, 687.7, 695.7, 0.2),
    "L21": (0, 514.5, 565.9, 706.5, 714.5, 0.2),
    "L22": (0, 514.5, 565.9, 725.2, 733.2, 0.2),
    "FZ1": (0, 514.5, 565.9, 762.7, 770.7, 0.2),
    "FZ2": (0, 514.5, 565.9, 781.5, 789.5, 0.2),
    "L23": (0, 514.5, 565.9, 800.0, 808.0, 0.2),
    # --- 第2页 ---
    "L23_2": (1, 514.5, 565.9, 10.0, 27.6, 0.2),
    "FZ3": (1, 514.5, 565.9, 43.0, 51.0, 0.2),
    "L24": (1, 514.5, 565.9, 61.7, 69.7, 0.2),
    # --- 第2页签章 ---
    "agent_name": (1, 80.8, 184.8, 103.4, 111.4, 2.0),
    "agent_id": (1, 112.8, 264.8, 113.0, 121.0, 2.0),
    "receiver": (1, 355.7, 463.7, 103.4, 111.4, 2.0),
    "receive_date": (1, 355.7, 483.7, 122.6, 130.6, 2.0),
    # --- 第3页 ---
    "A201_R1C1": (2, 224.0, 288.0, 104.0, 112.0, 0.2),
    "A201_R1C2": (2, 288.0, 352.0, 104.0, 112.0, 0.2),
    "A201_R1C3": (2, 352.0, 416.0, 104.0, 112.0, 0.2),
    "A201_R1C4": (2, 416.0, 480.0, 104.0, 112.0, 0.2),
    "A201_R1C5": (2, 480.0, 566.0, 104.0, 112.0, 0.2),
    "A201_R2C1": (2, 224.0, 288.0, 120.7, 128.7, 0.2),
    "A201_R2C2": (2, 288.0, 352.0, 120.7, 128.7, 0.2),
    "A201_R2C3": (2, 352.0, 416.0, 120.7, 128.7, 0.2),
    "A201_R2C4": (2, 416.0, 480.0, 120.7, 128.7, 0.2),
    "A201_R2C5": (2, 480.0, 566.0, 120.7, 128.7, 0.2),
    "A201_R3C1": (2, 224.0, 288.0, 135.7, 143.7, 0.2),
    "A201_R3C2": (2, 288.0, 352.0, 135.7, 143.7, 0.2),
    "A201_R3C3": (2, 352.0, 416.0, 135.7, 143.7, 0.2),
    "A201_R3C4": (2, 416.0, 480.0, 135.7, 143.7, 0.2),
    "A201_R3C5": (2, 480.0, 566.0, 135.7, 143.7, 0.2),
}


def fill_and_render(pdf_bytes, values, dpi=300, fmt="png"):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    font_data = extract_font_data(doc)
    font_simsun = None
    if font_data:
        try:
            font_simsun = fitz.Font(fontbuffer=font_data)
            st.info(f"✅ 提取字体成功: {len(font_data)} bytes")
        except Exception as e:
            st.warning(f"⚠️ 提取字体失败: {e}")
    
    comma_font_path = make_comma_font()
    font_embedded = None
    if comma_font_path:
        try:
            font_embedded = fitz.Font(fontfile=comma_font_path)
            st.info("✅ 内嵌字体载入成功（含逗号/星号/中文）")
        except Exception as e:
            st.warning(f"⚠️ 内嵌字体载入失败: {e}")
    
    use_fallback = (font_simsun is None)
    if use_fallback:
        st.warning("⚠️ 使用备份字体 china-ss")
    
    try:
        for key, raw_value in values.items():
            if not raw_value or key not in FIELD_CFG:
                continue
            
            new_value = fmt_decimal(raw_value, key)
            if not new_value:
                continue
            
            page_num, x0, x1, y0, y1, right_margin = FIELD_CFG[key]
            page = doc[page_num]
            text = str(new_value)
            
            # 逐个清除检测到的旧文字（只清除文字区域，不碰边框线）
            INSET = 2.0
            for b in page.get_text("dict")["blocks"]:
                if "lines" not in b:
                    continue
                for line in b["lines"]:
                    for span in line["spans"]:
                        sb = span["bbox"]
                        if (sb[0] >= x0 - 2 and sb[2] <= x1 + 2 and
                            sb[1] >= y0 - 2 and sb[3] <= y1 + 2):
                            cl = max(sb[0] + INSET, x0 + INSET)
                            cr = min(sb[2] - INSET, x1 - INSET)
                            ct = max(sb[1] + INSET, y0 + INSET)
                            cb = min(sb[3] - INSET, y1 - INSET)
                            if cr > cl and cb > ct:
                                rect = fitz.Rect(cl, ct, cr, cb)
                                shape = page.new_shape()
                                shape.draw_rect(rect)
                                shape.finish(color=(1, 1, 1), fill=(1, 1, 1))
                                shape.commit()
            
            # 写入新文字
            origin_y = y0 + (y1 - y0) * 0.75
            
            if key in ["agent_name", "agent_id", "receiver", "receive_date"]:
                # 签章字段：用内嵌字体（含完整中文+*号）
                write_x = x0 + 2.0
                if use_fallback:
                    page.insert_text((write_x, origin_y), text, fontname="china-ss",
                                     fontsize=FONTSIZE, color=(0, 0, 0))
                else:
                    tw = fitz.TextWriter(page.rect)
                    # 优先使用内嵌字体（含星号），回退到提取字体
                    font_to_use = font_embedded if font_embedded else font_simsun
                    tw.append((write_x, origin_y), text, fontsize=FONTSIZE, font=font_to_use)
                    tw.write_text(page, color=(0, 0, 0))
            else:
                # 数值字段: 逐字符右对齐，逗号/小数点用内嵌字体
                total_width = 0
                for char in text:
                    if char == ',':
                        total_width += COMMA_W
                    elif char == '.':
                        total_width += DOT_W
                    else:
                        total_width += CHAR_W
                
                current_x = (x1 - right_margin) - total_width
                
                if use_fallback:
                    for char in text:
                        page.insert_text((current_x, origin_y), char, fontname="china-ss",
                                         fontsize=FONTSIZE, color=(0, 0, 0))
                        if char == ',':
                            current_x += COMMA_W
                        elif char == '.':
                            current_x += DOT_W
                        else:
                            current_x += CHAR_W
                else:
                    tw = fitz.TextWriter(page.rect)
                    for char in text:
                        if char in [',', '.', '*'] and font_embedded is not None:
                            tw.append((current_x, origin_y), char, fontsize=FONTSIZE, font=font_embedded)
                        else:
                            tw.append((current_x, origin_y), char, fontsize=FONTSIZE, font=font_simsun)
                        
                        if char == ',':
                            current_x += COMMA_W
                        elif char == '.':
                            current_x += DOT_W
                        else:
                            current_x += CHAR_W
                    tw.write_text(page, color=(0, 0, 0))
        
        images = []
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            if fmt == "png":
                img_bytes = pix.tobytes("png")
            else:
                img_bytes = pix.tobytes("jpeg")
            images.append(img_bytes)
        
        return images
    finally:
        doc.close()


def main():
    st.title("PDF智能填表系统 v14.4")
    st.markdown("TextWriter + fitz.Font | **零外部依赖+星号修复+边框保护版** | PNG输出")

    st.header("1️⃣ 上传PDF模板")
    uploaded_file = st.file_uploader("选择PDF文件", type=["pdf"])

    if uploaded_file is None:
        st.info("👆 请先上传PDF模板文件")
        return

    pdf_bytes = uploaded_file.getvalue()

    st.header("2️⃣ 填写字段数据")
    st.caption("留空表示不修改。数字自动添加两位小数")

    values = {}

    with st.expander("👥 第1行 - 从业人数", expanded=True):
        cols = st.columns(5)
        for col, label, key in zip(cols, ["Q1季初","Q1季末","Q2季初","Q2季末","Q3季末(总)"],
                                   ["eq1s","eq1e","eq2s","eq2e","eq3e"]):
            with col: values[key] = st.text_input(label, value="", key=key)

    with st.expander("💰 第2行 - 资产总额", expanded=True):
        cols = st.columns(5)
        for col, label, key in zip(cols, ["Q1季初","Q1季末","Q2季初","Q2季末","Q3季末(总)"],
                                   ["aq1s","aq1e","aq2s","aq2e","aq3e"]):
            with col: values[key] = st.text_input(label, value="", key=key)

    with st.expander("📊 第3-16行（预缴税款计算）", expanded=True):
        col1, col2 = st.columns(2)
        left = [("L1","3 营业收入"),("L2","4 营业成本"),("L3","5 利润总额"),("L4","6 特定业务"),
                ("L5","7 不征税收入"),("L6","8 减：资产加速折旧"),("L7","9 减：免税收入")]
        right = [("L8","10 减：所得减免"),("L9","11 减：所得减免其他"),("L10","12 实际利润额"),
                 ("L11","13 税率(25%)"),("L12","14 应纳所得税额"),("L13","15 减免所得税额"),("L13_1","15.1 减免明细")]
        with col1:
            for k,l in left: values[k] = st.text_input(l, value="", key=k)
        with col2:
            for k,l in right: values[k] = st.text_input(l, value="", key=k)

    with st.expander("📋 第16-25行"):
        col1, col2 = st.columns(2)
        left2 = [("L14","16 本期预缴"),("L15","17 减免其他"),("L16","18 本期应补(退)"),
                 ("L17","19 总机构本期"),("L18","20 总机构分摊"),("L19","21 财政集中")]
        right2 = [("L20","22 分支机构"),("L21","23 分摊比例"),("L22","24 税率"),
                  ("FZ1","附1 中央级收入"),("FZ2","附2 地方级收入"),("L23","25 减免地方")]
        with col1:
            for k,l in left2: values[k] = st.text_input(l, value="", key=k)
        with col2:
            for k,l in right2: values[k] = st.text_input(l, value="", key=k)

    with st.expander("📄 第2页 - 补充字段"):
        col1, col2, col3 = st.columns(3)
        with col1: values["L23_2"] = st.text_input("23.2 本年累计应减免", value="", key="L23_2")
        with col2: values["FZ3"] = st.text_input("FZ3 地方级收入实际应纳税额", value="", key="FZ3")
        with col3: values["L24"] = st.text_input("24 实际应补(退)所得税额", value="", key="L24")

    with st.expander("✏️ 第2页 - 签章信息"):
        c1, c2 = st.columns(2)
        with c1:
            values["agent_name"] = st.text_input("经办人", value="", key="agent_name")
            values["agent_id"] = st.text_input("经办人身份证号", value="", key="agent_id")
        with c2:
            values["receiver"] = st.text_input("受理人", value="", key="receiver")
            values["receive_date"] = st.text_input("受理日期", value="", key="receive_date")

    with st.expander("📎 第3页 - A201020 资产加速折旧附表"):
        st.caption("行1: 加速折旧")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: values["A201_R1C1"] = st.text_input("R1 账载折旧", value="", key="A201_R1C1")
        with c2: values["A201_R1C2"] = st.text_input("R1 税收一般规定", value="", key="A201_R1C2")
        with c3: values["A201_R1C3"] = st.text_input("R1 加速政策计算", value="", key="A201_R1C3")
        with c4: values["A201_R1C4"] = st.text_input("R1 纳税调减金额", value="", key="A201_R1C4")
        with c5: values["A201_R1C5"] = st.text_input("R1 加速优惠金额", value="", key="A201_R1C5")
        st.caption("行2: 一次性扣除")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: values["A201_R2C1"] = st.text_input("R2 账载折旧", value="", key="A201_R2C1")
        with c2: values["A201_R2C2"] = st.text_input("R2 税收一般规定", value="", key="A201_R2C2")
        with c3: values["A201_R2C3"] = st.text_input("R2 加速政策计算", value="", key="A201_R2C3")
        with c4: values["A201_R2C4"] = st.text_input("R2 纳税调减金额", value="", key="A201_R2C4")
        with c5: values["A201_R2C5"] = st.text_input("R2 加速优惠金额", value="", key="A201_R2C5")
        st.caption("行3: 合计")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: values["A201_R3C1"] = st.text_input("R3 账载折旧", value="", key="A201_R3C1")
        with c2: values["A201_R3C2"] = st.text_input("R3 税收一般规定", value="", key="A201_R3C2")
        with c3: values["A201_R3C3"] = st.text_input("R3 加速政策计算", value="", key="A201_R3C3")
        with c4: values["A201_R3C4"] = st.text_input("R3 纳税调减金额", value="", key="A201_R3C4")
        with c5: values["A201_R3C5"] = st.text_input("R3 加速优惠金额", value="", key="A201_R3C5")

    st.header("3️⃣ 输出设置")
    c1, c2 = st.columns(2)
    with c1:
        output_fmt = st.radio("图片格式", ["PNG（无损推荐）", "JPG"], horizontal=True)
    with c2:
        dpi = st.select_slider("DPI", options=[150, 200, 300, 400], value=300)
    fmt = "png" if "PNG" in output_fmt else "jpeg"
    ext = "png" if fmt == "png" else "jpg"

    filled = {k: v.strip() for k, v in values.items() if v.strip()}
    if filled:
        st.success(f"✅ 已填写 {len(filled)} 个字段")
    else:
        st.info("💡 尚未填写任何字段")

    if st.button("🚀 生成图片", type="primary", disabled=len(filled)==0):
        with st.spinner("正在渲染..."):
            try:
                images = fill_and_render(pdf_bytes, filled, dpi=dpi, fmt=fmt)
                st.success(f"✅ 成功生成 {len(images)} 页图片（{dpi} DPI）")
                if len(images) == 1:
                    st.download_button(f"📥 下载图片（{ext.upper()}）", images[0],
                        f"filled_tax_form_page1.{ext}", mime=f"image/{ext}", use_container_width=True)
                else:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for i, img_data in enumerate(images):
                            zf.writestr(f"filled_tax_form_page{i+1}.{ext}", img_data)
                    zip_buf.seek(0)
                    st.download_button(f"📥 下载全部（ZIP）", zip_buf.getvalue(),
                        "filled_tax_form.zip", mime="application/zip", use_container_width=True)
                st.subheader("预览（第1页）")
                st.image(images[0], use_container_width=True)
            except Exception as e:
                st.error(f"❌ 生成失败: {e}")
                st.exception(e)

    st.markdown("---")
    st.markdown("<center>PDF智能填表系统 v14.4 | TextWriter + fitz.Font | 零外部依赖+星号修复+边框保护版</center>",
                unsafe_allow_html=True)


if __name__ == "__main__":
    main()
