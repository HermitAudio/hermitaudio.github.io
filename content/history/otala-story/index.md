+++
title = "The Story of the \"Otala\" Amplifier"
weight = 10
+++

*A story of the legendary 25W "Otala" amplifier as seen by Terje Sandstrøm, one
of the original designers.*

*Also see "[The People Involved](../people-involved)", for other designers and
people who contributed.*

## Introduction — the birth of the Otala amplifier

At an AES conference in 1973, [Dr. Matti Otala](http://www.tut.fi/~strategy/otala/)
presented a paper describing the design of a TIM-free audio amplifier. Present
at this conference was Svein Erik Børja, a Norwegian record and broadcasting
producer, and a great audio enthusiast. Svein Erik Børja was also one of the
greatest Golden Ears of his time — he was able to hear even the slightest
imperfection in an audio component. Having been dissatisfied with the sound of
the transistorized audio amplifiers of that day, he saw an opportunity here for
an audio amplifier of a new generation. Dr. Otala's talk about TIM also
explained the imperfections Svein Erik himself had noted in audio amplifiers.

He brought the paper to a friend of his who was running his own audio company
in Norway, Per Abrahamsen of Electrocompaniet. Per decided to give the
amplifier a try, to see what it might bring. They got the help of Nils Jørgen
Kjærnet at Nera in Oslo, who did the circuit board design and also, to my
knowledge, helped with the mechanical design of the amplifier.

They first made just a couple of amplifiers, but the sound was so good,
fulfilling all their expectations, that they decided to manufacture a series of
these amplifiers.

I started showing up at Electrocompaniet in the autumn of '74, and started to
work for Per in the spring of '75. One of my first jobs was to assemble the
first production series of the Otala amplifiers — a series of 10 amplifiers,
based on the same PCBs as the original two prototypes. During the summer I
started to look at the design, having been an electronics hobbyist since my
early teens. At that time neither Per nor I knew too much about high-end audio
design. That may well have been the factor behind the amplifier's success, as
it evolved over the next 5 years — we didn't know how it should be done, so we
worked it out from the framework given by the Otala design. We may well have
been the first audio amplifier designers of the new school — the TIM-free
designers, so to speak.

## The first Otala amplifier in "production"

The first amplifier series used the same printed circuit boards as the first
two prototypes. The PCB was a two-layer design — we called it a T-board, due to
its shape. The power transistors were mounted onto the cooling fin from both
sides of the stem of the "T". One layer of the board was a ground plane (Nera
being an RF company, that wasn't so strange). The ground plane caused us a lot
of mounting problems, because the component legs often got stripped when they
were squeezed through the holes in the PCB, causing small metal pieces from the
legs to curl up and make sweet, nearly invisible, shorts to the ground plane.
We did see smoke ...

In the autumn of '75 we started to get some attention, and a visit to Matti
Otala at the VTT (Valtion Teknillinen Tutkimuskeskus — the Technical Research
Centre of Finland), where he worked as a professor, became necessary. There
were several reasons:

We had named the amplifier "Otala-Lohstroh" after its inventors, and Otala
wanted to see the amplifier before he and Lohstroh could possibly allow us to
use his name.

At the same time, we started to experience problems with the amplifier — it
didn't quite meet the specifications stated in the AES paper. We needed his
help.

As it turned out, the amplifier was renamed to "The 2-channel audio power
amplifier", which was the name used for the rest of the amplifier's lifetime.
The relationship with Otala was established, and he visited the company on
several occasions. There is still good contact between Otala and
Electrocompaniet.

## Period of confusion, start of revision

It was early 1976. We started to invest in measuring equipment. The company
slowly turned into an audio amplifier company. It lost several of its old
customers; the loudspeakers it sold became more neglected. We worked from
early morning into the late nights. That would be the standard for the next
4 years. An amplifier brought to the US by Svein Erik <span id="amplifier-blow-up"></span>blew
up at the first listening test — embarrassing. We used the night to invent a
new type of [short-circuit protection](../protection-networks). We didn't want
anything that could affect the sound, so we simply used a high-impedance
circuit to sense the current through the output transistors, and then used a
relay to switch off the voltage to the output transistors. It worked!

We now understood that we had to start changing the amplifier. Svein Erik's
golden ears became even more important than ever before. We were in uncharted
territory; the old measurement methods could not be trusted. We had to rely on
our own methods, our own interpretations of the measurement results. Now began
an active period of theory, practice, and reading articles (Svein Erik provided
us with articles en masse, on all aspects of audio design and electronics
design in general).

We found that the compensation scheme used in the amplifier did not work
correctly. As we measured the bandwidth to be far less than specified, we
needed to redesign these filters. The amplifier used shunt feedback — one of
its really strong points — with an input lag compensation; at each of the
three amplifying stages, a lead-lag compensation was introduced. These did not
match the actual poles we measured. One of the reasons for this, we assumed,
was that the original design used "fresh," newly developed Philips
transistors. As both Otala and Lohstroh were working at the Philips labs in
Eindhoven at the time of the design, we assumed they had access to better
transistors. (I don't remember if Otala ever confirmed this to be the case.)

In any case, a period followed where we changed the compensation back and
forth with little or no effect. We did achieve something, but no significant
improvement.

We had also now measured the distortion (having just bought our first
distortion measurement set and spectrum analyzer). It wasn't pretty — the
distortion was far worse than specified, and performance was far below even
our own expectations. Again we didn't manage to do anything significant at
first.

A major change came first when we started to change the quiescent current of
the transistors. The design was in fact done as a mix between the old way and
the new way. Older design books teach that you should use low current in the
first stage due to noise. This was also done here (see "Noise Optimisation"
for an explanation of why that's not actually the case). We knew this theory
was wrong, so we increased the quiescent currents, which decreased the
resistance and again dramatically increased the bandwidth of the amplifier.
The distortion also went down — not dramatically, but enough to make a
significant change for the better in the sound (see "Transistors, Resistors,
Current and Distortion" — not yet written).

## The Audio Critic

At this time we started to hear rumours about a
[test](/magazine-articles/early-electrocompaniet/audio-critic-1976/) in
the US on our amplifier. One day our mail started to overflow, and a few days
later we got hold of the test ourselves. It was in a magazine named The Audio
Critic, and the test was fabulous. It started:

> Audio freaks - Eat your hearts out: This is the worlds best sounding
> amplifier, and One: You can't buy one in this country (The USA!), and Two:
> (not surprisingly - the only slightly negative in the whole test) It is too
> low powered to be counted.

Suddenly we had more requests for amplifiers than we could hope to handle.
Sales boomed. It was mid-'76.

We worked all summer, mornings, days, evenings, and into the night — Monday to
Friday, Saturday and Sunday. Per's wife, Anne, was not always too happy about
this. Per had two small boys, and she had to take care of them a lot by
herself. Anyway — her "sacrifice" made building these amplifiers possible, and
the nice thing about it all is that both Per's and Anne's sons now work at
Electrocompaniet! So in the end it turned out to be an investment in their own
future as well!

## The great change

Dr. Otala's theorems on TIM very often came out as an attack on high-feedback
amplifiers. Although it is correct that high-feedback amplifiers are more
prone to TIM than low-feedback amplifiers, there is no magic here. The main
goal of any amplifier is to reproduce the incoming music as perfectly as
possible, neither adding nor subtracting anything. TIM is just one type of
distortion; swapping one type for another doesn't help. It is true that some
types of distortion are harsher to the ear than others, but there still is a
balance to be struck. If 0.1% of TIM equals 1% of THD in audibility, then an
amplifier with 2% THD and 0.07% TIM will sound worse than an amplifier with 1%
THD and 0.1% TIM.

So, we realised that the balance of distortions was the essential factor to
consider. Not only do you have to balance THD against TIM, but also
low-frequency distortion against high-frequency distortion, frequency and
phase response against non-linear distortion in general, and so on.

This insight triggered the Great Change:

One night (it always happened at night!) we increased the feedback by 10 dB,
for a total of 30 dB of feedback. The sound improvement was staggering — and
contrary to common belief in our own community!

After this we only adjusted the amplifier slightly. It had found the form it
should have for its remaining life at Electrocompaniet. And the amplifier was
a success!

## On marketing

We did brochures! The [first one](../brochures#our-first-brochure) was rather
technical. The [second one](../brochures#our-second-brochure) was more
professional. The rest of the marketing was done by word-of-mouth. The fact
that the amplifier and the company were both hard to get hold of stimulated
the market. Here was a mysterious company with a killer amplifier. Audio
people love such stuff!

## Technical things

See the [schematics](/schematics/evolution-of-the-schematics/). Also see some
of the [calculations](/schematics/calculations/) that were done.
