# MemoNotes Music training software

I'm making my software available on GitHub (MIT license) to anyone interested in it.

I made a video demonstrating some of its features, I tried to show it as much as possible in the shortest possible time --> https://www.youtube.com/watch?v=MkNB0SxQOi8

[![Alt text](https://img.youtube.com/vi/MkNB0SxQOi8/0.jpg)](https://www.youtube.com/watch?v=MkNB0SxQOi8)

**EDIT:** *There is a slight delay in the audio, but this is due to the Camtasia software that records the screen. This happens every time I record with it, maybe it's something I haven't set up correctly on it. But MemoNotes software has great accuracy and precision in terms of time. I didn't use the traditional Timer object, which has a maximum precision of 50 milliseconds and is also quite affected by other processes, which can cause a huge delay. Instead of Timer I used the System.Diagnostics.Stopwatch class which achieves a much higher precision/resolution and has a higher priority over other Windows processes. So rest assured that the audio in the software plays with great precision, with no lag, even when the pace is very fast.*

Well, it's the first time I've published something on GitHub, so I hope I uploaded the files correctly, and that everything works fine when downloading the project --> https://github.com/LazieWouters/MemoNotes---Music-training-software

I didn't use this code for a long time, but I tested it now in Visual Studio 2022 Preview and it ran, I took the opportunity to upgrade to the 4.8 framework.

I made this software many years ago, more or less between 2009 and 2012, I'm not sure, but it took about 2 years between the ideation and completion of the project. It was made for my personal use in learning music, I didn't make it for the purpose of meeting other people's needs.

I programmed it in VB .NET because my initialization in programming was creating macros for Access and Excel with VB for Applications. So, at the time, I felt more familiar with the VB .NET language than with C#. After that I migrated my studies to C# to develop different projects.

At the time I didn't have much knowledge of good programming practices either, so you will definitely see a lot of things in the code that could have been better implemented. But the important thing is that the executable works and is extremely compact, despite the project having a large Mb size due to the PNG images used and audio files for the musical notes and chords. That's why I'm going to urge you to avoid comments like "this part of the code would be better implemented if it were that way", "here you jerry-rigged tremendously in the code" and things like that, I request this because today I know I could have programmed in a better way with the knowledge I have today (10 years later) and also because I'm not going to change the code anymore. But if you still want to give constructive criticism that's fine then.

The software has many things in the Brazilian Portuguese language in its interface, which is my language of origin because as I said, I didn't make it for sale.

Some system features:

* Training for memorization and recognition of notes on the musical staff (score);

* Tone recognition using FFT (Fast Fourier Transform), so you can use the software as training to tune your vocals or your acoustic guitar;

* Musical scale training (all scales used in all cultures are available);

* MIDI connection for musical keyboards and digital pianos;

* Guitar chord training;

* Rhythm training with musical notes in a score that displays infinite combinations of notes and rests depending on selected options;

* Tonal interval training;

* Chord sequence training;

* and more things;

I hope this project can be useful to someone else.


![alt text](https://preview.redd.it/g5my75ujbp681.png?width=1920&format=png&auto=webp&s=7276f9694bb461ffba2efc44f524ef4f124bd1eb)

![alt text](https://preview.redd.it/ic7igdtnbp681.png?width=1920&format=png&auto=webp&s=1336b22d5d51e801223e1efe131bab1eb1581b61)

![alt text](https://preview.redd.it/he8ue5oqbp681.png?width=1920&format=png&auto=webp&s=510d2e1189499cc5b287952559931e2965724a3b)

![alt text](https://preview.redd.it/0mvks9subp681.png?width=1920&format=png&auto=webp&s=dbf28258b7a1c61cab8fa90da60179b7c9839d73)

![alt text](https://preview.redd.it/ayjwby6xbp681.png?width=1920&format=png&auto=webp&s=d19241f0f158478aae4630e59b86622d3ca74367)

![alt text](https://preview.redd.it/i3vxlm7zbp681.png?width=1920&format=png&auto=webp&s=a3dffe6a883e94a6d393c1a50431298ae74d2e8a)

![alt text](https://preview.redd.it/cp96gca1cp681.png?width=1920&format=png&auto=webp&s=753ff2a94a975456c5dee925d5b985490f2b7347)

![alt text](https://preview.redd.it/etfum0p2cp681.png?width=1920&format=png&auto=webp&s=1c666146d917faea57dc92bef6bcbc4ef8f1765c)

![alt text](https://preview.redd.it/y2kmosr6cp681.png?width=1920&format=png&auto=webp&s=4ad48fff9f3e96e42f8acd41bb797fbb5e7b3df6)

![alt text](https://preview.redd.it/16w5frzacp681.png?width=1920&format=png&auto=webp&s=d0b80ee57fee772a69fdb3df1f19a2f588a067ea)
