library(readxl)
library(openxlsx)
library(stringr)
###导入额外信息
addinfo<-read.table('add_info.txt',encoding='UTF-8')

###今天的麦克表单数据
fileslist<-list.files(paste0(getwd(),'/mike'))
#储存数据
frame_mike<-data.frame()
for (i in 1:length(fileslist)){
  #循环导入excel文档
  tempoframe<-read_excel(paste0(getwd(),'/mike/',fileslist[i]),skip=1)
  #锁定"想报名的线下活动"这一列
  tempoframe2<-tempoframe[,which(str_detect(names(tempoframe),'想报名的线下活动'))]
  if (ncol(tempoframe2)==0){
    next()
  }
  names(tempoframe2)<-'报名场次'
  #提取出"想报名的线下活动"中包含对应城市的数据
  tempoframe<-tempoframe[str_detect(tempoframe2$`报名场次`,as.character(addinfo$V1[1])),]
  tempoframe<-cbind(tempoframe$`手机`,tempoframe$`邮箱`)
  frame_mike<-rbind(frame_mike,tempoframe)
}

###今天新加入的数据
## 需要保证两点:数据列中仅有一列名中含有"手机",第一行不要有其它的东西
fileslist2<-list.files(paste0(getwd(),'/add'))
frame_add<-data.frame()
if (length(fileslist2)!=0){
for (i in 1:length(fileslist2)){
  tempoframe<-read_excel(paste0(getwd(),'/add/',fileslist2[i]))
  #搜索列名中包含"手机"和"邮箱"的列
  tel<-tempoframe[,which(str_detect(names(tempoframe),'手机'))]
  email<-tempoframe[,which(str_detect(names(tempoframe),'邮箱'))]
  #防止文件内邮箱是空的
  if (ncol(email)==0){
    email<-rep(NA,length(tel))
  }
  tempoframe<-cbind(tel,email)
  names(tempoframe)<-c('手机号','邮箱')
  frame_add<-rbind(frame_add,tempoframe)
}
}

###历史数据导入
fileslist3<-list.files(paste0(getwd(),'/history'))
frame_history_mail<-data.frame()
frame_history_tel<-data.frame()
#将history文件夹里的所有文件分成短信和邮件两类
filelist3_mail<-fileslist3[str_detect(fileslist3,'mail')]
filelist3_tel<-fileslist3[str_detect(fileslist3,'message')]

if (length(filelist3_mail)!=0){
for (i in 1:length(filelist3_mail)){
  #批量导入历史数据中的短信
  tempoframe<-read_excel(paste0(getwd(),'/history/',filelist3_mail[i]),col_names = FALSE,sheet=2)
  frame_history_mail<-rbind(frame_history_mail,tempoframe)
}}

if (length(filelist3_tel)!=0){
for (i in 1:length(filelist3_tel)){
  #批量导入历史数据中的电话
  tempoframe<-read_excel(paste0(getwd(),'/history/',filelist3_tel[i]),col_names = FALSE,sheet=2)
  frame_history_tel<-rbind(frame_history_tel,tempoframe)
}}

###开始查重,今天发所有的
sendtel<-c(as.character(frame_mike$V1),as.character(frame_add$`手机号`))
history_tel<-as.character(frame_history_tel$...1)
sendtel<-sendtel[is.na(match(sendtel,history_tel))]
##添加自己的联系方式,方便自己收到信息,以作测试
sendtel<-c(sendtel,'13922250231','13507436010')
# 稍微清洗一下
sendtel<-str_replace_all(sendtel,' ','')
#除去文本两侧的空格
sendtel<-str_trim(sendtel)
#自身去重
sendtel2<-unique(sendtel)
sendtel2<-sendtel2[!is.na(sendtel2)]


##保存电话序列
write.xlsx(sendtel2,paste0(getwd(),'/sendtel.xlsx'),na='',row.names = FALSE,col.names = FALSE,sep=',')

##保存邮件序列
sendemail<-c(as.character(frame_mike$V2),as.character(frame_add$`邮箱`))
history_mail<-as.character(frame_history_mail$...1)
#和历史数据去重
sendemail<-sendemail[is.na(match(sendemail,history_mail))]

##添加自己的联系方式,方便自己收到信息,以作测试
sendemail<-c(sendemail,'919086038@qq.com','1148235075@qq.com')
##自身去重
sendemail2<-unique(sendemail)
sendemail2<-sendemail2[!is.na(sendemail2)]
#保存邮件发送的数据
write.table(sendemail2,paste0(getwd(),'/sendmail.csv'),na='',row.names = FALSE,col.names = FALSE,sep=',',quote=FALSE)


####下面为上传onedrive的文件准备了
today<-strftime(Sys.Date(),'%m%d')
##今日电话发送列表存档
write.xlsx(list('【待补充今日日期】'=sendtel2,'【待补充历史日期】'=sendtel2),paste0(getwd(),'/MPB_event_message_',today,'_',as.character(addinfo$V1[2]),'_',length(sendtel2),'#.xlsx'),na='',row.names = FALSE,col.names = FALSE,sep=',')
##今日邮件发送列表存档
write.xlsx(list('【待补充今日日期】'=sendemail2,'【待补充历史日期】'=sendemail2),paste0(getwd(),'/MPB_event_mail_',today,'_',as.character(addinfo$V1[2]),'_',length(sendemail2),'#.xlsx'),na='',row.names = FALSE,col.names = FALSE,sep=',')


