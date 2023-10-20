#streamlit==1.27.2
#SpeechRecognition==3.10.0
#pyaudio==0.2.13 #(cho SpeechRecognition lay micro )
#googletrans==4.0.0rc1 #(phien ban nay cho rieng py khi su dung googletrans, cac pban khac hay gay loi)
#gTTS==2.4.0


#https://talkenvi-b5vypm7itcecxnkuvne7h9.streamlit.app/ 
#la url app moi talkenvi
import streamlit as st
import speech_recognition as sr 
from audio_recorder_streamlit import audio_recorder #pip install audio-recorder-streamlit
from googletrans import Translator 
from gtts import gTTS, gTTSError   
from io import BytesIO  
#import numpy as np
#import soundfile as sf
#import sounddevice as sd


def speech_to_text(lang):
    #red="#FF0000" , blue="#0000FF" , yellow="#FFFF00"     
    if lang=="vi-VN":
        audio_bytesa = audio_recorder(text='A.(Say in Vi - Nói bằng tiếng Việt):',recording_color="#FFFF00",neutral_color="#FF0000",icon_size="2x",energy_threshold=(-1.0,1.0),pause_threshold=3.0)
        if audio_bytesa:
            with open('thua.wav','wb') as fa:
                fa.write(audio_bytesa)
                r = sr.Recognizer()
                with sr.AudioFile('thua.wav') as source:
                    audioa = r.record(source)  # read the entire audio file
                try:
                    texta = r.recognize_google(audioa, language=lang)
                    return texta
                except sr.UnknownValueError:
                    st.write("Không thể xác định giọng nói.")
                except sr.RequestError as e:
                    print(f"Lỗi: {e}")
    elif lang=="en_US":
        audio_bytesb = audio_recorder(text='B.(Say in En - Nói bằng tiếng Anh):',recording_color="#FFFF00",neutral_color="#0000FF",icon_size="2x",energy_threshold=(-1.0,1.0),pause_threshold=3.0)
        if audio_bytesb:
            with open('thub.wav','wb') as fb:
                fb.write(audio_bytesb)
            #st.audio(audio_bytesb, format="audio/wav")
                r = sr.Recognizer()
                with sr.AudioFile('thub.wav') as source:
                    audiob = r.record(source)  # read the entire audio file
                try:
                    textb = r.recognize_google(audiob, language=lang)
                    return textb
                except sr.UnknownValueError:
                    st.write("Không thể xác định giọng nói.")
                except sr.RequestError as e:
                    print(f"Lỗi: {e}")
    
def textsrc_to_textdest(l_text, lang_src,lang_dest):
    translator = Translator()
    translation = translator.translate(l_text, src=lang_src, dest=lang_dest)
    #st.write(translation.text)
    return translation.text

def text_to_speech(text, lang='vi'):
    try:
        tts = gTTS(text, lang=lang)
        data_io = BytesIO()
        tts.write_to_fp(data_io)
        data_io.seek(0)
        #chuyen data_io sang nparray am thanh nho sf roi choi nparray nho sd
        #data, samplerate = sf.read(data_io)
        #sd.play(data, samplerate)
        #sd.wait()        
        return data_io
    except gTTSError as err:
        st.error(err)
    

#######################################################
st.subheader(":blue[Trò chuyện tiếng Việt (có thông dịch) với...]")
#vaichon = st.radio(":green[Select one of options to say:]", 
#                [":orange[Vietnamse]", ":blue[English]",":green[Danish]",":orange[German]",":yellow[Taiwan]",":blue[Japanese]",":red[Korean]","CANCEL"],
#                index=7,horizontal=True ) 
noi_voi = st.selectbox("Chon nguoi Noi voi:", 
                ("English (en)","Spanish (es)","Taiwan (zh-TW)","Danish (da)","German (de)","Japanese (ja)","Korean (ko)"),index=0,label_visibility="hidden")
sub1='('
sub2=')'


mtext1="**A** Nói tiếng Việt (Say in Vietnamese):"
audio_bytes1 = audio_recorder(text=mtext1,recording_color="#FFFF00",neutral_color="#FF0000",icon_size="2x",energy_threshold=(-1.0,1.0),pause_threshold=3.0)
if audio_bytes1:
    txt1=''
    txt2=''
    txt_translated1=''
    txt_translated2=''    
    lang='vi'
    lang_src='vi'
    lang_dest='en'

    with open('thu1.wav','wb') as f1:
        f1.write(audio_bytes1)
        r = sr.Recognizer()
        with sr.AudioFile('thu1.wav') as source1:
            audio1 = r.record(source1)  # read the entire audio file
        try:
            text1 = r.recognize_google(audio1, language=lang)
            st.write(text1)
        except sr.UnknownValueError:
            text1=''
            st.write("")
        except sr.RequestError as e:
            text1=''
            print(f"Lỗi: {e}")
            st.write("")
    if text1 !='':
        txt_translated1 = textsrc_to_textdest(text1, lang_src, lang_dest)
        st.write(txt_translated1)
    if txt_translated1 !='':
        audio_io = text_to_speech(txt_translated1, lang_dest)
        st.audio(audio_io, format="audio/wav",start_time=0)

st.write("---")
mtext2='**B** Nói tiếng '+noi_voi+' (Say in '+noi_voi+'):'
audio_bytes2 = audio_recorder(text=mtext2,recording_color="#FFFF00",neutral_color="#0000FF",icon_size="2x",energy_threshold=(-1.0,1.0),pause_threshold=3.0)
if audio_bytes2:
    txt1=''
    txt2=''
    txt_translated1=''
    txt_translated2=''    

    lang='en'
    lang_src='en'
    lang_dest='vi'

    with open('thu2.wav','wb') as f2:
        f2.write(audio_bytes2)
        r = sr.Recognizer()
        with sr.AudioFile('thu2.wav') as source2:
            audio2 = r.record(source2)  # read the entire audio file
        try:
            text2 = r.recognize_google(audio2, language=lang)
            st.write(text2)
        except sr.UnknownValueError:
            text2=''
            st.write("")
        except sr.RequestError as e:
            text2=''
            print(f"Lỗi: {e}")
            st.write("")
    if text2 !='':
        txt_translated2 = textsrc_to_textdest(text2, lang_src, lang_dest)
        st.write(txt_translated2)
    if txt_translated2 !='' :
        audio_io = text_to_speech(txt_translated2, lang_dest)
        st.audio(audio_io, format="audio/wav",start_time=0)

exit()
st.write("---")
audio_bytes=None
if vaichon == ":orange[Vietnamse]":
    mtext="Vietnamse: click Mic to say..."
    lang="vi"
    lang_src=''
    lang_dest='vi'
    audio_bytes = audio_recorder(text=mtext,recording_color="#FFFF00",neutral_color="#0000FF",icon_size="2x",energy_threshold=(-1.0,1.0),pause_threshold=3.0)

elif vaichon==":blue[B. Nói tiếng Anh (Say En):sunflower:]":
    st.write(":blue[B. Nói tiếng Anh (Say En):sunflower:]")
    lang="en"
    lang_src='en'
    lang_dest='vi'
    lang_dest=''
elif vaichon==":red[C. Nói tiếng Trung (Say Zh):sunflower:]":
    st.write(":red[C. Nói tiếng Trung (Say Zh):sunflower:]")
    lang="zh"
    lang_src='zh'
    lang_dest='vi'
    lang_dest2='en'
else:    
    #st.write("")
    lang=""
    lang_src=''
    lang_dest=''
    lang_dest=''
#B1: ghi am giong noi va chuyen thanh text

if audio_bytes and lang != '':
    with open('thua.wav','wb') as f:
        f.write(audio_bytes)
        r = sr.Recognizer()
        with sr.AudioFile('thu.wav') as source:
            audio = r.record(source)  # read the entire audio file
        try:
            text_from_audio = r.recognize_google(audio, language=lang)
        except sr.UnknownValueError:
            st.write("Không thể xác định giọng nói.")
        except sr.RequestError as e:
            print(f"Lỗi: {e}")

    #B2: dich sang text En hoac Vi
    if text_from_audio is not None:
        txt_translated = textsrc_to_textdest(text_from_audio, lang_src, lang_dest)
        st.write(txt_translated)
    if txt_translated is not None:
        audio_io = text_to_speech(txt_translated, lang_dest)
        st.audio(audio_io, format="audio/wav",start_time=0)
