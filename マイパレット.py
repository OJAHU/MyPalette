from programs import System, GUI
import json

class Application():
    def __init__(self):
        json_path = r"C:\Myfiles\Settings.json"
        
        with open(json_path, "r", encoding="utf-8") as f:
            settings = json.load(f)
            
        self.mymemo = settings["mymemo"]
        self.mytask = settings["mytask"]
        self.vba_name = settings["vba_name"]
        self.token = settings["token"]
        self.credentials = settings["credentials"]

        self.boot()
    
    def boot(self):
        while True:
            system = System.System(self.mymemo,
                                   self.mytask,
                                   self.vba_name,
                                   self.token,
                                   self.credentials)
            
            gui = GUI.Display(system.tasks, 
                              system.subject_lst,
                              system.time_lst,
                              system.until_today,
                              system.until_tommorow,
                              system.msgs)
        
            if gui.status["RESTART"] == False:
                break

if __name__ == "__main__":
    Application()