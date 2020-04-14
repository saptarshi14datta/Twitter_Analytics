# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Twitter_GUI.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!
import tweepy
import time
import json
import os
import threading
import pandas as pd
import threading
import sys
import csv
import random
import win32com.client, pythoncom
from Twitter_Access import consumer_key,consumer_secret,access_token,access_token_secret
from PyQt5 import QtCore, QtGui, QtWidgets


#Global variables

global api, count, myStream, ui, directory

count = 0
targets = []

class MyStreamListener(tweepy.StreamListener):
    def on_status(self, status):
        global directory
        
        if hasattr(status, "retweeted_status"):  # Check if Retweet
            try:
                print(status.retweeted_status.extended_tweet["full_text"].encode("utf-8"))
                text = status.retweeted_status.extended_tweet["full_text"].encode("utf-8")
                with open('OutputStreaming.txt', 'a') as f:
                    writer = csv.writer(f)
                    writer.writerow([status.author.screen_name, status.created_at, status.retweeted_status.extended_tweet["full_text"].encode("utf-8")])
            except AttributeError:
                print(status.retweeted_status.text.encode("utf-8"))
                text = status.retweeted_status.text.encode("utf-8")
                with open('OutputStreaming.txt', 'a') as f:
                    writer = csv.writer(f)
                    writer.writerow([status.author.screen_name, status.created_at, status.retweeted_status.text.encode("utf-8")])
        else:
            try:
                print(status.extended_tweet["full_text"].encode("utf-8"))
                text = status.extended_tweet["full_text"].encode("utf-8")
                with open('OutputStreaming.txt', 'a') as f:
                    writer = csv.writer(f)
                    writer.writerow([status.author.screen_name, status.created_at, status.extended_tweet['full_text'].encode("utf-8")])

            except AttributeError:
                print(status.text.encode("utf-8"))
                text = status.text.encode("utf-8")
                with open('OutputStreaming.txt', 'a') as f:
                    writer = csv.writer(f)
                    writer.writerow([status.author.screen_name, status.created_at, status.text.encode("utf-8")])
        
        reply_tweet(status.id, text)
    
    
    
    def on_error(self, status_code):
        if status_code == 420:
            #returning False in on_error disconnects the stream
            return False
    
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(557, 345)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.btn_quit = QtWidgets.QPushButton(self.centralwidget)
        self.btn_quit.setGeometry(QtCore.QRect(460, 240, 75, 23))
        self.btn_quit.setObjectName("btn_quit")
        self.btn_run = QtWidgets.QPushButton(self.centralwidget)
        self.btn_run.setGeometry(QtCore.QRect(460, 210, 75, 23))
        self.btn_run.setObjectName("btn_run")
        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox.setGeometry(QtCore.QRect(10, 10, 411, 151))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.groupBox.setFont(font)
        self.groupBox.setObjectName("groupBox")
        self.label = QtWidgets.QLabel(self.groupBox)
        self.label.setGeometry(QtCore.QRect(10, 30, 171, 16))
        self.label.setObjectName("label")
        self.lbl_target_user = QtWidgets.QPlainTextEdit(self.groupBox)
        self.lbl_target_user.setGeometry(QtCore.QRect(190, 20, 211, 31))
        self.lbl_target_user.setObjectName("lbl_target_user")
        self.label_2 = QtWidgets.QLabel(self.groupBox)
        self.label_2.setGeometry(QtCore.QRect(10, 70, 171, 16))
        self.label_2.setObjectName("label_2")
        self.lbl_no_slaves = QtWidgets.QPlainTextEdit(self.groupBox)
        self.lbl_no_slaves.setGeometry(QtCore.QRect(190, 60, 211, 31))
        self.lbl_no_slaves.setObjectName("lbl_no_slaves")
        self.label_3 = QtWidgets.QLabel(self.groupBox)
        self.label_3.setGeometry(QtCore.QRect(10, 110, 171, 16))
        self.label_3.setObjectName("label_3")
        self.lbl_media_dir = QtWidgets.QPlainTextEdit(self.groupBox)
        self.lbl_media_dir.setGeometry(QtCore.QRect(190, 100, 211, 31))
        self.lbl_media_dir.setObjectName("lbl_media_dir")
        self.groupBox_2 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_2.setGeometry(QtCore.QRect(430, 10, 120, 151))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.groupBox_2.setFont(font)
        self.groupBox_2.setObjectName("groupBox_2")
        self.chk_set_target = QtWidgets.QCheckBox(self.groupBox_2)
        self.chk_set_target.setGeometry(QtCore.QRect(10, 30, 101, 17))
        self.chk_set_target.setChecked(True)
        self.chk_set_target.setObjectName("chk_set_target")
        self.chk_get_slaves = QtWidgets.QCheckBox(self.groupBox_2)
        self.chk_get_slaves.setGeometry(QtCore.QRect(10, 60, 101, 17))
        self.chk_get_slaves.setObjectName("chk_get_slaves")
        self.chk_analyze_data = QtWidgets.QCheckBox(self.groupBox_2)
        self.chk_analyze_data.setGeometry(QtCore.QRect(10, 90, 101, 17))
        self.chk_analyze_data.setObjectName("chk_analyze_data")
        self.chk_take_action = QtWidgets.QCheckBox(self.groupBox_2)
        self.chk_take_action.setGeometry(QtCore.QRect(10, 120, 101, 17))
        self.chk_take_action.setObjectName("chk_take_action")
        self.lbl_log = QtWidgets.QTextBrowser(self.centralwidget)
        self.lbl_log.setGeometry(QtCore.QRect(10, 170, 411, 131))
        self.lbl_log.setObjectName("lbl_log")
        self.btn_abt = QtWidgets.QPushButton(self.centralwidget)
        self.btn_abt.setGeometry(QtCore.QRect(460, 270, 75, 23))
        self.btn_abt.setObjectName("btn_abt")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 557, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        MainWindow.setTabOrder(self.lbl_target_user, self.lbl_no_slaves)
        MainWindow.setTabOrder(self.lbl_no_slaves, self.lbl_media_dir)
        MainWindow.setTabOrder(self.lbl_media_dir, self.btn_run)
        MainWindow.setTabOrder(self.btn_run, self.btn_quit)
        MainWindow.setTabOrder(self.btn_quit, self.btn_abt)
        MainWindow.setTabOrder(self.btn_abt, self.chk_set_target)
        MainWindow.setTabOrder(self.chk_set_target, self.chk_get_slaves)
        MainWindow.setTabOrder(self.chk_get_slaves, self.chk_analyze_data)
        MainWindow.setTabOrder(self.chk_analyze_data, self.chk_take_action)
        MainWindow.setTabOrder(self.chk_take_action, self.lbl_log)
        
        self.btn_run.clicked.connect(self.execute_program)
        self.btn_quit.clicked.connect(self.exit_program)
        self.btn_abt.clicked.connect(self.about_program)
    
    def about_program(self):
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        myWord = win32com.client.DispatchEx('Word.Application')
        print(os.getcwd())
        script_dir = os.path.dirname(os.path.realpath('__file__'))
        rel_path = 'Help_File.docx'
        abs_file_path = os.path.join(script_dir, rel_path)
        print(abs_file_path)
        #wordfile = r"C:\Users\joy\Documents\Python Scripts\Twitter\Help_File.docx"
        myDoc = myWord.Documents.Open(abs_file_path, False, False, True)      
        
        
    def exit_program(self):
        sys.exit()
    
    def print_log(self, text):
        word = '<span style=\" color: #009900;\">%s</span>' % text
        self.lbl_log.append(text)
        
    def get_inputs(self):
        if self.chk_set_target.isChecked():
            print("module_set_target")
            self.module_set_target = True
        else:
            self.module_set_target = False
        if self.chk_get_slaves.isChecked():
            print("module_get_slaves")
            self.module_get_slaves = True
        else:
            self.module_get_slaves = False
        if self.chk_analyze_data.isChecked():
            print("module_analyze_data")
            self.module_analyze_data = True
        else:
            self.module_analyze_data = False
        if self.chk_take_action.isChecked():
            print("module_take_action")
            self.module_take_action = True
        else:
            self.module_take_action = False
        self.target_user = self.lbl_target_user.toPlainText()
        self.no_slaves = self.lbl_no_slaves.toPlainText()
        self.media_dir = self.lbl_media_dir.toPlainText()
        
        
        
    def execute_program(self):
        global api, myStream, target_subset, directory
        self.get_inputs()
        directory = self.media_dir
        
        self.print_log("Executing program")
             
        twitter_authentication()
        set_targets(self.target_user)
            
        if self.module_get_slaves == True:
            get_slaves(int(self.no_slaves),self.target_user) 
            if self.module_analyze_data == True:
                analyze_data()
        
        if self.module_take_action == True and self.module_get_slaves == False:
            x = threading.Thread(target = monitor_targets,args=(api.get_user(self.target_user).id_str,), daemon=True)
            x.start()
        elif self.module_take_action == True and self.module_get_slaves == True:
            for index,row in target_subset.iterrows():
                x = threading.Thread(target = monitor_targets,args=(str(row['User_ID']),), daemon=True)
                x.start()  
        

        print("Done")
        
    
        

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.btn_quit.setText(_translate("MainWindow", "Quit"))
        self.btn_run.setText(_translate("MainWindow", "Run"))
        self.groupBox.setTitle(_translate("MainWindow", "Input"))
        self.label.setText(_translate("MainWindow", "Target User (for monitoring)"))
        self.label_2.setText(_translate("MainWindow", "No. of Slaves (for monitoring)"))
        self.label_3.setText(_translate("MainWindow", "Media Files Directory"))
        self.groupBox_2.setTitle(_translate("MainWindow", "Modules to Run"))
        self.chk_set_target.setText(_translate("MainWindow", "Set Targets"))
        self.chk_get_slaves.setText(_translate("MainWindow", "Get Slaves"))
        self.chk_analyze_data.setText(_translate("MainWindow", "Analyze Data"))
        self.chk_take_action.setText(_translate("MainWindow", "Take Action"))
        self.btn_abt.setText(_translate("MainWindow", "About"))


def twitter_authentication():
    global api
    # Authenticate to Twitter using OAuth 1a
    auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
    auth.set_access_token(access_token, access_token_secret)
    api = tweepy.API(auth, wait_on_rate_limit=True, wait_on_rate_limit_notify=True, compression=True)

def set_targets(user):
    #The target twitter account(s)
    targets.append(user)

def get_slaves(max_slaves, user):
    global api,ui
    i=0
    j=0
    for user in targets:
        get_data(user,"master") 
        for page in tweepy.Cursor(api.followers_ids, screen_name=user).pages():
            for follower in page:
                try:
                    if j < max_slaves:
                        try:
                            get_data(follower,"slave")
                            j+=1
                        except:
                            pass
                    else:
                        return None
                except:
                    print("user not found")
                    pass
            i=i+1
            print('page ', i)

def get_data(user_x,category_x):
    global api, count, ui
    user_category = category_x
    user_id = api.get_user(user_x).id_str
    user_name = api.get_user(user_x).name
    user_location = api.get_user(user_x).location
    user_description = api.get_user(user_x).description
    user_url = api.get_user(user_x).url
    user_followers_count = api.get_user(user_x).followers_count
    user_friends_count = api.get_user(user_x).friends_count
    user_listed_count = api.get_user(user_x).listed_count
    user_statuses_count = api.get_user(user_x).statuses_count
    user_verified = api.get_user(user_x).verified
    
    data = {'Category':user_category,
             'User_ID':user_id,
             'User_Name':user_name,
             'User_Location':user_location,
             'User_Description':user_description,
             'User_Url':user_url,
             'User_Followers_Count':user_followers_count,
             'User_Friends_Count':user_friends_count,
             'User_Listed_Count':user_listed_count,
             'User_Status_Count':user_statuses_count,
             'User_Verified':user_verified       
        }
    
    print(user_name)
    ui.print_log("Gathering data for user: "+user_name)
    
    with open('followers.json', 'a', encoding="utf-8") as f:
        count +=1
        if count < 2 :
            json.dump(data,f)
        else:
            f.write(",")
            json.dump(data,f)            
    f.close

def analyze_data():
    global target_subset,ui
    print(os.getcwd())
    twitter_data = pd.read_json(r'C:\Users\joy\Documents\Python Scripts\Twitter\followers.json', lines=True)
    twitter_data = twitter_data.sort_values(by=['User_Followers_Count'],ascending=False)

    #target_subset = twitter_data[(twitter_data.User_Followers_Count >= 100) | (twitter_data.User_Friends_Count >= 100)]


    target_subset = twitter_data[(twitter_data.User_Followers_Count >= twitter_data["User_Followers_Count"].mean()) |
                     (twitter_data.User_Friends_Count >= twitter_data["User_Friends_Count"].mean()) |
                     (twitter_data.User_Listed_Count >= twitter_data["User_Listed_Count"].mean()) |
                     (twitter_data.User_Status_Count >= twitter_data["User_Status_Count"].mean())
                     ]  
    
    target_subset = target_subset.drop_duplicates(subset='User_ID', keep='first')


def monitor_targets(target_id):
    global myStream, api, ui
    print("Monitoring Targets")
    myStreamListener = MyStreamListener()
    myStream = tweepy.Stream(auth = api.auth, listener=myStreamListener,tweet_mode='extended')
    ui.print_log("Monitoring target: "+target_id)
    #myStream.filter(track=['python'], is_async=True)
    myStream.filter(follow=[target_id])
    with open('OutputStreaming.txt', 'w') as f:
        writer = csv.writer(f)
        writer.writerow(['Author', 'Date', 'Text'])


def reply_tweet(tweet_id, text):
    global api,ui,directory
    list1=["artless","bawdy","beslubbering","bootless","churlish","cockered","clouted","craven","currish","dankish","dissembling","droning","errant","fawning","fobbing","froward","frothy","gleeking","goatish","gorbellied","impertinent","infectious","jarring","loggerheaded","lumpish","mammering","mangled","mewling","paunchy","pribbling","puking","puny","quailing","rank","reeky","roguish","ruttish","saucy","spleeny","spongy","surly","tottering","unmuzzled","vain","venomed","villainous","warped","wayward","weedy","yeasty"]
    list2=["base-court","bat-fowling","beef-witted","beetle-headed","boil-brained","clapper-clawed","clay-brained","common-kissing","crook-pated","dismal-dreaming","dizzy-eyed","doghearted","dread-bolted","earth-vexing","elf-skinned","fat-kidneyed","fen-sucked","flap-mouthed","fly-bitten","folly-fallen","fool-born","full-gorged","guts-griping","half-faced","hasty-witted","hedge-born","hell-hated","idle-headed","ill-breeding","ill-nurtured","knotty-pated","milk-livered","motley-minded","onion-eyed","plume-plucked","pottle-deep","pox-marked","reeling-ripe","rough-hewn","rude-growing","rump-fed","shard-borne","sheep-biting","spur-galled","swag-bellied","tardy-gaited","tickle-brained","toad-spotted","urchin-snouted","weather-bitten"]
    list3=["apple-john","baggage","barnacle","bladder","boar-pig","bugbear","bum-bailey","canker-blossom","clack-dish","clotpole","coxcomb","codpiece","death-token","dewberry","flap-dragon","flax-wench","flirt-gill","foot-licker","fustilarian","giglet","gudgeon","haggard","harpy","hedge-pig","horn-beast","hugger-mugger","jolthead","lewdster","lout","maggot-pie","malt-worm","mammet","measle","minnow","miscreant","moldwarp","mumble-news","nut-hook","pigeon-egg","pignut","puttock","pumpion","ratsbane","scut","skainsmate","strumpet","varlet","vassal","whey-face","wagtail",]
    slang = "Thou "+list1[random.randint(0,int(len(list1)-1))]+list2[random.randint(0,int(len(list2)-1))]+list3[random.randint(0,int(len(list3)-1))]
    random_file=random.choice(os.listdir(directory))
    abs_file_path = os.path.join(directory, random_file)
    #api.update_with_media(abs_file_path,status="Thanks",in_reply_to_status_id=tweet_id,auto_populate_reply_metadata=True)
    ui.print_log("Responded to tweet id: "+str(tweet_id)+" which contained "+ str(text))
    ui.print_log("==================================================")



if __name__ == "__main__":
    global ui
    import sys
    app=QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    MainWindow.setWindowTitle("Twitter Analytics -- v2")
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())