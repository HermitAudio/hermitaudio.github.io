+++
title = "A Theory of Single Stages"
weight = 10
math = true
+++

## The general idea

When you design something, you're moving your ideas and abstract thoughts
into the real world. Using electronic components, they will never act
exactly equal to your ideas. So the best approach is to approximate their
behavior, and make them work as close to your ideas as possible. Further, if
you choose your models skillfully, you will get a behavior from the device
which very closely mimics what you're after.

Therefore, the sequence is: Model ⇒ Design ⇒ Measurement ⇒ Listening — and you
circle this sequence until you're happy :-)

The walkthrough below is rather detailed, to show the general outline of the
procedure. The other derivations are less rigorous.

## A model for a single stage

We model a single stage as a voltage-controlled current source. In our ideal
world this means we want an active device working as a transconductance
device. A transistor (and also a tube) is very close to this ideal. But one
has to further improve the circuitry around the active device, in order to
make it behave as closely as possible to this ideal.

What this means is that a perfect transistor, and thus a perfect stage, will
have infinite input impedance, no reverse coupling from output to input,
infinite output impedance, and a finite and constant (with respect to both
the signal and the environment) transfer conductance.

Real transistors are not quite as good as this, but we can improve on the
transistor in order to make it behave more like this ideal. Doing this will
normally make the stage perform better in all respects, but keep in mind that
all rules will turn back on you at a certain stage. There is no such thing as
a free lunch.

If such a stage is voltage-driven, we will reduce the nonlinearity from all
"leakages" back to the input, be it input impedance or reverse coupling. The
dominating nonlinearity will then be the transconductance, which is easy to
control.

A transistor in a common-emitter coupling is the starting point. It has, in
principle, the behavior described above. The following rules exist:

**Linear behavior:**

Transfer conductance, given by:

$$g_m = \frac{I_e}{V_T}$$

$I_e$ is the emitter DC current and $V_T$ is the voltage equivalent of
temperature, normally equal to 25 mV — the exact formula is $kT/q$, where $k$
is Boltzmann's constant, $q$ is the charge of an electron, and $T$ is the
absolute temperature in Kelvin.

The inverse of the transconductance is called the dynamic resistance, called
$r_e$.

Current amplification: $H_{FE} = I_c/I_b$, derived for a particular current
$h_{fe} = i_c/i_b$, which applies for small-signal currents around a
quiescent point $I_c$.

Input impedance: $r_{in} = h_{fe} \cdot r_e$

The dominating nonlinear mechanism lies in the transconductance. Since this
is a single stage (not a differential stage) it will generate a smooth series
of harmonics (if stimulated with a pure sinusoid).

If the input signal is given as $x$ (where $x$ can be e.g. $\sin(\omega t)$),
then the output $y$ will be:

$$y = a_1 x + a_2 x^2 + a_3 x^3 + \cdots$$

And $x$ is a relative parameter which must always obey $|x| < 1$ in order for
the series to converge. If $|x| < 1$ then it follows that $|y| < 1$. The
output current is $i_e$ and the output parameter is then $i_e/I_e$. The input
signal generating a current of $i_e$ is $u_{in}$, from the formula
$i_e = u_{in} \cdot g_m$. It then follows from the transconductance formula
that the input parameter we seek is $u_{in}/V_T$.

So, given the input $x$, defined as $u_{in}/V_T$, and the output $y$, defined
as $i_e/I_e$ — what do these things mean?

If $u_{in}/V_T$ exceeds 1, then the varying part of the current exceeds the
quiescent current $I_e$, and the stage is clipping. When the stage is
clipping, our formula breaks down. If we want to find the distortion when the
stage is clipping, we'll have to resort to Fourier analysis.

To make this into a practical case: assume a transistor running at a current
of 1mA. The $g_m$ is then 1/25 siemens (the inverse of ohm, the unit for
transconductance, although "mhos" — ohm reversed — is also used). An input
signal of 1 mV will then generate an output current of 1/25 mA = 40 µA. But
now note: this is the first-order approximation. As can be seen from the
series expansion above, we also have second- and third-order components. The
first-order coefficient should be pretty close to $g_m$, but what are the
other two coefficients?

The real equation relating input voltage to output current is
$I_e = I_s \cdot \exp\left(\frac{U_{be}}{V_T} - 1\right)$, called the
Ebers-Moll equation. $I_s$ is the "leakage" current, but don't bother about
it — we'll soon enough get rid of it. We are interested in finding the
equation for the behavior around the quiescent point. We do this by adding
small deviations $i_e$ and $u_{be}$ to the equation above. Resolving this, we
get the much simpler equation: $i_e/I_e = \exp(u_{be}/V_T) - 1$, and its
inverse: $u_{be}/V_T = \ln(1 + i_e/I_e)$. These equations are called the
signal equations, and will be used to get the coefficients for the series
expansion above. We'll first make a series expansion of the first equation,
and we get:

$$\frac{i_e}{I_E} = \frac{u_{be}}{V_T} + \frac{(u_{be}/V_T)^2}{2} + \frac{(u_{be}/V_T)^3}{6} + \cdots$$

## The efficiency parameter

We now introduce the efficiency parameter $n_i$. The point of introducing
this parameter is to generate simpler formulas for calculating the
distortion, and to gain a better understanding of how the distortion and
other transistor parameters are coupled.

The parameter is defined as:

$$n_i = \frac{i_e}{I_e}$$

For a bipolar stage without local feedback, as the stage discussed above,
$n_i$ is equivalent to $u_{be}/V_T$, where we only consider the linear part
of the series expansion. The equation above can then be written as:

$$\frac{i_e}{I_E} = n_i + \frac{n_i^2}{2} + \frac{n_i^3}{6} + \cdots$$

The second-order distortion is defined as the second-order term divided by
the first-order term, and the third-order distortion in the same manner:

$$\text{2nd} = \frac{n_i}{2} \qquad \text{3rd} = \frac{n_i^2}{6}$$

If the input signal is a sinusoid, then the following equations hold (ask via
[the about page](/about/) if you'd like the proof), where $\text{2ndh}$ is
the second-order harmonic distortion, and $\text{3rdh}$ is the third-order
harmonic distortion:

$$\text{2ndh} = \frac{\text{2nd}}{2} \qquad \text{3rdh} = \frac{\text{3rd}}{4}$$

Putting it together:

$$\text{2ndh} = \frac{n_i}{4} \qquad \text{3rdh} = \frac{n_i^2}{24}$$

...and remember, $n_i = u_{in}/V_T$.

## FET stages

It is interesting to do the same exercise for FET transistors. The result is
a simpler series, with only first- and second-order components. By inserting
the efficiency parameter, one gets the following equation for the FET's
second-harmonic distortion:

$$\text{2ndh} = \frac{n_i}{8}$$

Half the amount of the bipolar transistor. Some people have argued that the
FET is a much more linear device than the bipolar. This equation shows that
to be only a partial truth — there is only a 6dB improvement.

## Local feedback

It is well known that local series current feedback (read: inserting an
emitter resistor) reduces the distortion. The feedback factor can be written
as:

$$D = 1 + g_m R_e = 1 + \frac{I_e R_e}{V_T}$$

and the resulting efficiency/distortion equations are then:

$$\text{2nd} = \frac{n_i}{2D} \qquad \text{3rd} = \frac{n_i^2}{3D}$$

or for the harmonic distortion:

$$\text{2ndh} = \frac{n_i}{4D} \qquad \text{3rdh} = \frac{n_i^2}{12D}$$

Note that the efficiency parameter $n_i$ is still defined as $i_e/I_e$, but
the input version is now $n_i = u_{in}/V_{th}$, where $V_{th} = V_T D$.

## Differential stages

There is a similar set of equations for the differential pair.

If the stage is completely in balance (which of course rarely happens), all
second-order components will be cancelled.

The no-feedback solution will have a basic transconductance of
$g_m = I_e/(2V_T)$, which gives $n_i = u_{in}/2V_T$.

The corresponding distortion is then equal to the single stage, except for a
mismatch parameter $m$ and a common-mode signal factor $c$:

$$\text{3rd} = \frac{n_i^2}{3} \qquad \text{2nd} = \frac{(m+c) \, n_i}{2}$$

If the current source feeding the emitters has infinite output impedance, $c$
approaches 0. The formula for $c$ is $c = i_k/(2i_e)$. More information on
this is in my AES paper (not yet ported to this site). There used to be an
ActiveX Single Stage Calculator here as well — it's retired along with the
rest of the site's old ActiveX controls.

## Summary

| Stage type                 | 2nd harmonic       | 3rd harmonic         |
|-----------------------------|--------------------|-----------------------|
| Bipolar single stage        | $n_i/(4D)$         | $n_i^2/(12D)$         |
| FET single stage            | $n_i/(8D)$         | Ideally zero          |
| Bipolar differential stage  | $(m+c) \, n_i/(4D)$ | $n_i^2/(12D)$         |

## Original derivation notes (kuriositet)

The treatment above uses a simplified "efficiency parameter" shortcut. My
original 1978 handwritten working notes derive the same distortion results
the long way — a full circuit-level analysis in terms of the actual resistor
values ($R_C$, $R_F$, $R_E$, $R_B$, $R_S$), for both a simple stage and the
differential pair, including two appendices (a Taylor series expansion of
$\ln(1+x)$, and inversion of a cubic series). That derivation is not
transcribed onto this page — only scanned here as the original artifact:

- [Nonlinear Distortion in Simple Transistor Stages (PDF, scanned, 1978)](nonlinear-distortion-in-simple-stages-1978.pdf)
- [Nonlinear Distortion in Simple Transistor Stages, Part II — The Differential Pair (PDF, scanned, 1978)](nonlinear-distortion-in-differential-stages-1978.pdf)
