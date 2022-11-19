clc;
clear;
close all;
home;

filename = 'Registration_Details.xls';
%[num,txt] = xlsread(filename);
txt = readcell(filename);

len=height(txt);

blankimage = imread('certificate.jpeg');

text = cell(len,4);

for i=1:len
    for j= 2:5
        text(i,j-1)=txt(i,j);
    end
end

for i=2:len
        text_all=[text(i,1) text(i,2) text(i,3) text(i,4)];
        
        position = [380 280;315 395; 605 490; 160 490];
        
        RGB = insertText(blankimage,position,text_all,'FontSize',28,'BoxOpacity',0);
        
        figure;
        imshow(RGB);
        y=i-1;
        filename=['Certificate_Topic_' num2str(y)];
        lastf=[filename '.png'];
        saveas(gcf,lastf);
end    