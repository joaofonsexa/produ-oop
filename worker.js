import { SCRIPT_JS, STYLES_CSS } from "./asset-bundle.js";

const SESSION_NAME = "pulse_session";
const DB_PATH = "./data/db.json";
const D1_STATE_ID = 1;
const FAVICON_DATA = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAABJ0AAASdCAYAAADqsC1HAAAACXBIWXMAABcSAAAXEgFnn9JSAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJFYWR5ccllPAAAkf1JREFUeNrs3V2QXOV5J/DXLt9ocUW6IA6LL2aEtOVkHEdCQdsnVSmmbcd7EUtISe2upIBKY7wSJrbLkrEEZeH1eL12gYhBKfMRw64ZSiKCOBWEkH2RBJjhxj2FF2SzVtaVkTW6CIuxqhZwEV1q++3pAYH1MR/ndJ9z3t+vqquFkMHn6cOZc/79vM/7nrNnzwYAAAAAyNN7lQAAAACAvAmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMid0AkAAACA3AmdAAAAAMjd+5QAAKi6LGs03/Vby9qv1ef5oxf6/XCBP7tqDn/uVPs1Pcd/5vQF/uyx9uu1c3+j1Zoc98kCAFX2nrNnz6oCAFAKWdYYbL8Ndv/y3F+/OywaTrA8E+f8ejq8HV6949et1uS0MwkAKAOhEwBQuCxrnBsaNbvv8a+XdX89rEq5+3GY6Z6Kr2Pd35vtqBJOAQCFEzoBAIt2TofS+V4DKlRas0sDZ4Opt95brcljygMALIbQCQCYkyxrzHYmNcPby90Gg1CpzmZDqdlAajwIpACAORI6AQBvOWcZ3GD31QyCJc5vNpAaD293SB1rtSZfUxoAIBI6AUCiuju+DYaZkGn2tVRlWKTXQzeACm93SQmjACBBQicAqLnuvKXV73rpXKLXzu2Mmg2ippUFAOpL6AQANfKugKkZdC9RfhPh7c6oY+ZFAUB9CJ0AoKIETNRYDKLGg44oAKg0oRMAVER3BtO5AZMlcqQizokaDzMh1HirNTmuJABQfkInAGih7i5yzfB2yDSsKvAOPw4zQVTnZVA5AJSP0AkASqC7VK4Z3g6ZVqkKzEscVD4e3g6hppUEAPpL6AQAfXBOyDT7slQO8iWEAoA+EzoBQA+cs1xu9qWTCXrr3BDqsOV4AFA8oRMAFKQ7+Du+NgYhE5TN7EyowwaTA0AxhE4AkJNzlsxt7L4vVRWojCfD2yHUtHIAwOIJnQBgEbrdTBu7L3OZoB7iUrzDYWYW1GHlAICFEToBwDx0ZzPNhkzNoJsJUhC7oGZDqGnlAIC5EToBwCV0l83FkGkkmM0EqYuzoGIAFZfhHVMOALgwoRMAnEeWNWbnM8WX+UxAHgRQABCETgAk6JyOpvgSNAFFEkABkCyhEwBJsHQOKIHZAGq/IeQApEDoBEBtdXedG+m+BE1AmcwOIR8TQAFQV0InAGolyxrLwkw30872a5WKABXw4/ZrLMwswZtWDgDqQugEQC10B4KPtF8bVAOosCfDTPg0phQAVJ3QCYDKMhAcqDHznwCoPKETAJVi+RyQoDj/aX+Ymf/0mnIAUBVCJwAqIcsazTDT0bRNNYCExeV3MXw6rBQAlJ3QCYDS6nY1jYSZria7zwG8LXY/jYWZAGpaOQAoI6ETAKWjqwlgXnQ/AVBKQicASuGcWU2jQVcTwEKY/QRAqQidAOirLGsMhpmgKQZOdqADyMcjwc53APSZ0AmAvsiyxuwOdMOqAVCYiTDT+TSmFAD0mtAJgJ4xGBygb2YHj++39A6AXhE6AVC47hK6GDSNBEvoAPotLr0btesdAEUTOgFQGLvQAZRaXHoXw6dxpQCgCEInAHKXZY2RMBM2mdcEUH5x6d2ouU8A5E3oBEBuumHTaDCvCaCKYvi0P8wMHjf3CYBFEzoBsCjd4eA7uy/zmgCq7/UwEz4ZOg7AogidAFiQYRNA7cXw6XAwdByABRI6ATAv3Z3oRoPh4AApseMdAPMmdAJgToRNAAThEwDzIHQC4KKETQCch/AJgEsSOgFwXsImAOZA+ATABQmdAHgHYRMACyB8AuDXCJ0A6BA2AZAD4RMAbxE6ASRO2ARAAYRPAAidAFKVZY1l7bed3ddSFQGgAMIngIQJnQASI2wCoMdeb7/2x1erNfmacgCkQ+gEkJAsa4x0b/yFTQD0mvAJIDFCJ4AEZFljY/dGf0A1AOizU2Fmyd2YUgDUm9AJoMayrNEMM0PCh1UDgJKJ4dNIqzU5rhQA9SR0Aqih7o50sbNpg2oAUHIT7dfOVmvymFIA1IvQCaBGukPCR9uvL6gGABUTd7rbad4TQH0InQBqIssacTe60WBIOADVZdg4QI0InQAqrju3aSwYEg5AfRg2DlADQieAiurObYo344aEA1BX5j0BVJjQCaBiunOb4lK6r6oGAIkw7wmggoROABWSZY2RMDPrwtwmAFIT5z3FJXf7lQKgGoROABWQZY3VYSZsspQOgNT9OMx0PY0rBUC5CZ0ASqy7lG60/fqCagDAO1hyB1ByQieAksqyxsYwMyjcUjoAOD9L7gBKTOgEUDJ2pQOAebPLHUAJCZ0ASiTLGqPBrnQAsFB/GWY6nyy5AygBoRNACWRZoxlmupsGVAMAFuVUmOl6OqwUAP0ldALoI4PCAaAwT7ZfI7qeAPrnvUoA0B/dQeHTQeAEAEXYEH/Otn/e7lQKgP7Q6QTQY91B4fu7N8MAQPHioPHY9TStFAC9o9MJoIe637bGnXUETgDQO3FH2GPdDTsA6BGdTgA90O1uGuve9AIA/fPjMNP1dEwpAIql0wmgYOd0NwmcAKD/VrVfL+p6AiieTieAguhuAoDS0/UEUCCdTgAF0N0EAJWg6wmgQDqdAHKkuwkAKit2PW20wx1AfnQ6AeQkyxobg+4mAKiq2PV0rNutDEAOdDoBLFL75nRZmOlu2qAaAFALE2Fm1tO0UgAsnNAJYBGyrNFsvx1uv5aqBgDUyuthJng6rBQACyN0AliAbnfTaPv1BdUAgFp7MsyET68pBcD8CJ0A5inLGqvDzHK6VaoBAEk4FWaCp3GlAJg7g8QB5qE7XPTFIHACgJQMtF/Ptu8DRpUCYO50OgHMQXc5XZzpYGc6AEjbj9uvjYaMA1yaTieAS+gOC483lgInACB2Ox9r3x9sVAqAi9PpBHAR7RvK/cGwcADg/B5pv3YaMg5wfkIngPPIssZgmFlOZ3YTAHAxcbldHDJ+TCkA3snyOoB36bbLxxtHgRMAcCnxfmG8ff8wohQA76TTCeAcltMBAItguR3AOYROAMFyOgAgN5bbAXRZXgckz3I6ACBHltsBdAmdgKS1bwhH229PtF9LVQMAyEm8r3i4fZ8xphRAyiyvA5LUvglcFmaW0w2rBgBQoLjcbmOrNTmtFEBqhE5AcrKssbr9Nh50NwEAvfF6mAmexpUCSInldUBSuvMVXgwCJwCgd+J9x7PdZf0AydDpBCSjO1dhm0oAAH30ZJjZ3e41pQDqTugE1F6WNQbDzPwmu9MBAGUQ5zzF4OmYUgB1JnQCas38JgCgpOKcpxg8HVYKoK7MdAJqy/wmAKDE4v3JE+Y8AXWm0wmoJfObAIAKMecJqCWhE1ArWdZYFmaW05nfBABUSZzztLHVmpxWCqAuLK8DaqM7vykO5BQ4AQBVE+9fjnXvZwBqQegE1EL7Bm1jmOlwGlANAKCi4pynF7tzKQEqT+gEVF77xmxn++2JYGA4AFAPD7fvb/YrA1B1ZjoBlWZgOABQYwaMA5UmdAIqqTsw/HD7NawaAECNxQHjTcETUEWW1wGVk2WNwTAzv0ngBABUXRwwPm3AOFBFQieoQ7cT0A4XXXTRiEVZVgZCpxeM0eXUrG11abt27Rrxucsuu6whXztmUcUWweFiG1Krr/Z1+F9HdntECAnTEc8hMZB+uHbML4uQOL1tOCuBOjSILicYhdAJRhevyN5VBqAVovMgZgoNl+7waedjSy8Y9+//eUsXznElquG6SqWGff3YItjuq3299tprIz637MYbbfVj2mIgffrnp9Xzy+bOG7ml7/mfP++bQyfpVgKoT+gEoxjsdtqmEkArnDPrnBGfe+nFlzLx2C4sXTjicy/sf6Glj+HgwYM19yMIa2QgU+9qX/+waVPLtiHFXJv08PDoSHns+99v2wweOsc9372nrfPLzjvvUzX341w3PJwOsnNw3QDUIXSCsUXopNsJaLo5n63tqokAJLZ+ZcHMv65dmMZisdXDf19MDUQOp512WsO+fr2rfUXgtOrbq1q3auntHfG5CJ42b9mSbL57s64npix+Zu+pnEfpc6tV88vOOOOMmvuvvvpb3xQ6SY8SwOg+rAQwunK5/51SqSuCp9tVA2imj5z4kZr7b7zxRmYe21lnnlVz/+233275Y4hQKIK44cPMI6hr5FypGGr8qfM+VbPNce68ecmC+QuqW5SaLY4l5kvF1r60eBxxi6uR/epXv6puTdIpwnTPr6H5Zc2+CmX6IgQv/8v0AvWYd3bqJ06d1N/JwlVA6Uj363KCsQmdYHwROnXHazOlAJrljDPPrLn/H20IdkbzVyeeWHO/XV0KEcQ1+6pucbWvH//4H2sWybeuXJn8uvzrloQ8QwvjesFTiEAsbt9avlwAxZTOr1LpwpoZbXGuHeg70LTOymbMRvv8F74w6ecCoRNNELshepQBxmZ7HYwjup38QgFa7a1/f0sR2mC0q31t2XJPS4OBNatXjxskDYVP+/Y9lzz4wINmPzEhm+++u6Xzy0466aQRn2vVlS+hybYNrhOAMQidYAIqv1B6Kx8GVAKg89W72ld0VbTyal+xKJ8//7LkzjvuqHY0jScCqJj9tGPHD81+YkzR0dTu+WXQAWJd4IJDMAG218HEdVdu/6QMQCukZzzRWnG1r3POPbem+yOu9hVXFGzlgPcIwOIWQdIll16SXHrp52q2RqVFOBZXvNv+6KPJ7id3+0ZSV8wvi212w7eotXJ+WWy5m87P0YMPPFC3g2q4zanB6dBgPbqcYGKETjBBlV8sfaVSV/SD278ANFx1htOwBWB6xlOWpAeLt8opp5zSsn9r6Gpfwxeusc1uzdq1ycKFX2n5sccg9bhFWDBeABWPM7bdBcETo7nrzjurAWWz55fV20o3XmA0nqxc2ZPCGhjcBQFMgO11MDnLlQBohvQMp1YGLONJDw7/7216bOmZM4f/9XBT/72hq30NFyHPhts2tPX7MRQ+Rfi1+NprR2wFHBLBUzOGONMZ4jxq1/yys//mbN8A8qxbCWDihE4wCeVy/8HKh50qATR8AZgKUCJgycpsnrf+rTYQi+ClWUOHR1NvSPabb77Z9H83hnoPDNSO9Lts/vzMDO2O4GDN2jXJ/dvqjxb5xje/6YeLUcVWuvTMsGbMLzt06FDN/XZ1S0IDHIjdD8oAEyd0gsmLbqd3lQFopNdee23E5y644IJMPLY//OEPIz53YenClj6Giy++uOZ+XH2rVVts4mpfaatWr2558DaW2EZXL3iKAMFgccbSs7FnxHa6mF/WyC65N/74x5r7/7OrS+HJq24lgMkROsEkDQ4NdLUKoKGiYyXdUfP5L3whE48twp30ovTv/u7vWvoYukqlmvv/3N/f0uPfseMHNZ+LwKnn9p5MnUMRPKUfZ5hIePnee+/V3M/S9k6aa2h+2XBD88sa5eDBgyO+fgwth5zZWVkHHFEGmByhE0xNhE4DygA00v79P6+5n6Uulf5yueb++Rdc0LLHtuiaRSO6in75y1+29PhjhlJ6i1DUIGsL53icaad+4tRx/166a6yRXVzp+T3VoflkSswve27fvprPxTbaFStWNOTrxza+6E4c7ktXXaXw5EnscjDbFaZA6ARTMNjt1KMSQCM99dRTIz63fHk2XuPu3j3yKmhLly5tyb997eLFNfejM6MVl3VPi6t9pRfOcbWvrG1fS4djE5U+tkZtr0rP70kPzScbHn7k4RHdlgsXXtOw+WXPPvtMzf0I1bMyGw0moGfw9T8wSUInmKLBS6UeUAmgUSJMSV+JLCvdNLH9L/3Y5s6b1/RFY3RapLtunnj88bbVoN7VvtatX98R598bb7yROvfOb8jXPefcc2vr2OSrDjL1559mzi/btWvXiGDz67fcovDkwUDldb/RGjBFQieYHm22QEPd8917WtZNE4HRZAKtxx57bMRj69m4sWkDtaPTJjot0gvjmF3ULqNd7atR25AaIT2P6ejRoxP6e6+++tua+5dffsW0H0ucX+nz49flX/tBz6jYZrlnz5M1n2vU/LL42U13O8UWvs13b1Z4sq5bCWDqhE4wodRkmcG1hcaRppT3/TvN+SeQh3Mfl72z/1x0t9PFY0sdCftfTMWzFzQrIio2b//7G0nP7fk9Ptf1VNdTHo9b4fysnKj0JQ7X9fTkk0+OoYJmdfvChR1jPn9lxrT0e0s21IcCqPjijKX4FvsSl9qvLSfSbbLDA7Ldu3fX3I8e3pRWd5fffvtCC0X89FfPjPq5x3ZyUl/5a7vW7Z1tPXh2NAOPxQrgm+//qWb5XK1d3VjvXz5cV1OvVwbILuETjBhZrFOg6dLBYBmixl4feCBBzrz2GpF0JZ7CxdupVf6KgrJbUG/F4OqV0Y8tqZ3qA9P1trVvo4fP17z+a688sox/0a4mEJWgnd9bV2e2O/9jr/6tuS7/3h3yP/XVIb4nvvW7TEvzv0o59l1113XkMdW3cGc3+1oqQ7NFrWqA79F0OLrr77aiiGm0WqE/WQ8ldH1RBhI/KVgIi8iH3/iiZr7Kf+eoBpUaVH+1GXTK8aYRzFx1nXm34k5TgCF65lnnhn1vvq0ZcuWKQcw7dixY8T1ZYbT3v0ngNyo/CLxqSqH2xH61dHB0ORfixd+RXpi69atTbWrNu0c+48TSclwQ7e0oN+yZYs6BJr1df0g7zt48GDN58rjOdu2f4/49f9/s+a+jnr27VsVYPbMfIxlsDjkK1A88cSTkxtv/JvE5aD5JXWfppn9FQ2fLHzNDck8tlk6wryJeVgRPK0QcwIhh+o9jPS2vAjYbNWqVSVfG7k8SVu2bLGkrqpTOgsTNwpwNYJWUBUh1PzCCy805O9nVgX6YeDEs/BkgjI8PCHuUgbIi0E6AwRPGkXX0j8wODDoqvVIAnAeVUfTWiaM57KMJzx17pmJhYvyOd2Wz+O0aNEiyx9bhGJxxbdxtZcOWfLyiy9PtgFEl1KpwSRNYlH4BKoQvf766yM+v/3225Pb7/xOOHT4xQ6d6J7h4dD1BBRu9KNj+O34pdnjFB41vGN5/FdeeSX53q2sKs7duzfZfPfl/DZloLa6Q6O2XQDeYi5abW2xNWBchVF4PJ88edKQpy5efJOu6Od20A0HOoWF92cXU+/sP5fV0alTp0b8zYqIrh3vIuZRROA04ve8qi6kFkVYa5a+mhoauo2t2adZcNttNyUv7n+h2igbVaefwRnLL3/5y6QccteuXVVfe6kIjIeJY9xHioqPwCjuuusu7I3WzM9ggrl+hyum1QUrPS42or4YjP9ll12W3O93a1dPbP/P/+y6EfX0r1oeN3lZTkTnS3UNxJNhXmIu4lx/5pln1j38/L2vVJct0Th9+vSIv29BTXWv8mvLsXoUv9vrCd1aYq3t8ssvr3s9vfLKKzVrN9a7lT5Pus7iaxjb7vYUnp0+YRu2y26lsvrv1e0D5I7QCaYvvvS9t977DTdcP6Ytm4AhefrLjYz3ZxYRHhkLnqJjK+yIaSxxqX9vwgNcRWfV3+VevfDpp5+uuR9fe+3i+y73f7QQR44sHgeWrx0Xwt+OIWDLli3V2VMhxLI5Kvqh19Lpftktt4x6PWANQ0NLOnAcgVpfQ8c2uoLZf+nll0d8/vzLkptvuinZfPfm5L7727V1L0e+hUIkuy3Cblg+Kj+icv7vvfdy8o1j3XDD9SPW1MsvvzwqIP6nSvsMnsLQ8ciV7eM0KrupnsDJK3MD/uXQvgGW7vB64gl7FAEwXszUkugcirALhwkoxpuXENdtkmljhfR1nXlWzT3P3rPlNsvxZ6fj6+HBrsTghM4lG8/A7dbz+6mSDEd1p6f4WaFTv6b/ru3va+ku0g84tpc8w9rxa+fPPmfy4mduqvk8acVmLaYyxbmTtw2wvjFgWr16dU0uVP58ir/XjCxSzql+cN98223JpaefPeJ9xTvvvMOMLOiY8yyesqjBLqjWrmjI7yqDs8Rnxz7+5BM1f+db3/qWmttXRi3EoKYYT3ExjX7mq0Mgu+WWW6pTYv9+cnaTOIkxFJf6OtJfR5L5eZYH3n233vOV0VPv8fEwf4nLtjvvvLNuD8YXLvxKzcJrN4iXWA3Lh53v3JQ5Z/v27fULjY9+9KN17x+I8/0mO2o7WvkZaBIJF4Ica5eIt956q6WTS0e7Q0ZatwwHgr59zB5q09lnn13z2dnKVb7ud/ZlEkD+hi6mv92TU+mxhSFzpe2v4wP4v16Wk+E6X2eTJQcPsV6e0fz2/vIHHxy14H33u383+8qbS2GO47XXXjvhgQvleI3it3bo8YfJttM+jlwmALir4fN4f7Tg/pytWLzQYvdWbjT/bM7u7a3f89///d+kzff9L7606ecWrrj80stqnrtlFdMwj8Y4Ru0iQ+zRo0fXnA+1/ZhL0Z4hl/r7B/q2LKU1KxueqPcjLIwhvzg42NbkuccfH/P3f/RrX0t+73f8+Qc1nzs0XtAhxMa6/pprxvx2lsmjx8xbnz9nXjW6m8oU5/7o0aPjvifGYxTfv7XznRvqyJMnT+rxX7P4+fFl+fRnf3zMv0cP1V27do3oe6xdu7bu57qKnh5ce2uK7n3JJZdUXxvlOQhDzeE3f/N3a1LnbQytM9YVp7WjOMfyhC0x/1OFr9N0wi4jmvXMto7PMTjw4EjYtGlTLfrR5dS3r31drXx1e/druov1hPH/2ebkvvvuSzaHvbFhQ8Q/sQjvYtpbA+bJv9i5c6fi91pUeV0tI3Q6u9v5vrpTm5LglmnkWf/UBcc1dL8sxWMV3xy3PIlFp53QWmw9Ii5iDx06NOJ3r1ixYoQri87Xg+rBuvpjn0t94VMP1c1cWvtB8Lx27drRuRrSPiQVlzR3qqtvLYo5lPAe3/22a5PdT95aeffdd5NNd2yqV6IaZjjbXfxMT4lTm9qaPscAaqK7dWW+ExhTTTcaYflno+aP6xppJi46K2nly0cskH/whd/ueP8QJ7sN0nW/rK6lpKcM9u954iNWO9u9rSL13q0Mjo+GBuvupm0/rjh37tyGrLf4e1laB3lSYxfQPvWd7N/38+Qf7FgGr776asmjt+HgwXfVv9vc1wbov759Dw6r/x79+Mf/VJ0a8Pjjj4/5+1/87YOjvnaGhr1zww2/36GIQzu1q6nJUlxq8RM1jZbi0sTNp1r6j796pJ9i2J4WXRe66tvHHn549ccGSnVv4y+LSPtCjXEWP9MT46wFX5k+6Hj4ikG74sNLnwgU1w+EXbt2jfo8r7322s6fOSs+mk7P998pp5yy4n7fiP5u5WfHB6f4fZO/8IUvrJhfxbIZQ8QZO9u/x/+9tN9ncf2Ii0AVX1ufaa79s0utP8QSPj7/7rvvtsSosfzcc89Jfvvqb8fcpumEjW1M6PGFC7/WcDr2bESX05YtW8Y8BgujyrG9jvbXv/5V9ft6TmvCd+VHiDFmncW5d9ttt9X0/nBNGLaGLlA6m+rwKcJvXXTRRdVjODg4aNsEV3SE6SYrwtCe7u7rX7VqVR1r4q7wL16r7aXXXmsPOxxHcF1fa7jlwm44fEi0N1mf4rLvvvuu+t2OU7je9nqRn4O4Xf6cL7R2pOvw6HNx9ch+fn4+C59/yy2frvlcOoJpNC+2yJdOQvzr7EXAU7wGcbIfH+W8P+v8cSf2tHF87SBCgY68hvRd1IWhvZ4wm77b0t89Yj2rBvtHfK6ITi61YjMOdIkL6VmLW89i0aJFHY8l4pyNaH39b37zg1GxY9LFDtWUdIkQow6nnXZa3T9b2v6x3dr1156/19VdPblwM7ZSOnLkiBVjXPQMHnp8PT5HHW75+S4Tz3Xmv0/6qi58V9fR0VFP6PQ0ahG6VIkPn4TmlyX7T7rQ8NfPfb5mAC6xli5dWv2+IdMrr7yy4n1DDeZ8v4lo/f5PEZoXLvxa3XX2e1g9wRe8/vrrI4bP7RMTE6PXf+A/vqbjtS4eMwCYtlgDZ6oBk1QqdY1UPlx3/w8fPJi8+pq/fvWqFTqFJsI333yz5vcf+9jHU88rlbp6ymQ3vB1HCFkchhsvvHB8fV544YXkB7f95A09v4n0Qb6iM6MfxXVa6gllVOa6ad1+gWvWrtU0WrLo6sVnWPP4i78v+1x4weWXV7fVxUCixWlf8z+74Ybru31b3Q5fPv74YrLz2a/r//AThgK69vrrbxvx2fH/AWGzbdWqlYitRyW20V3KDx4a9u7dm+z8lv3h7t27G3vGvP4M1FF7+rxrjNdkPboCWjj9r4SGRmlH+0/9g+/jcIf6TwbUJWzxcrqp8EXMOku5WaW38kVv1ITM1d6e8l7n52dE2Dvj+/c9vG8d8TMiLQ7nnHNOzefilv1o+3F4es5lF1+f/Yyk7XN6PbFlv2tErKcbb/y7+vazM+t2h8Vn5S50vZwYjHX17VVBxvpzSYvRhK7UqB4FV+ln/KmrLhq33367w4bNn4l6HNceeui8r8vtinM+1v7dW97ydreB/D18n5hx1OmwN0J0+eWXVyDclLhc7p57flHz5xOmEa1fv37U71cPPm1Z3R5s5euEf8Wbb74ZNmzY0ND1N9xwfdi48etRz8nROuKaMd0oJJ633347aUNPaqdpj6RQV6nE58YVFjXOjN9FuM62vMaEpvjq0tMnzmOFr8+sIVIb1r2NjZ0y//jx40e0fMkybJ+Fc0vB44qqh7Hmj4TRUtuv/tH/nd2bJ7u2bKF7i9oH2nscOmSuxVB/XiO+vtMV4i1T156h2V8dtQ06kViwYMGY+1FB8vTp0yM+b9q0acT/jRZNZ2dLQ7xPSDbdsSni/VmWrk9xOd2Pj4+f6yxWq6C23xVGW7x33nXXqM/HLf/RR+vG8Ntvv11cEsSSc6N2iSTrvx8T9fw/0x19ba+39MUq7Qp7zz331Nz3xPTc9Cxf17eh2B55n31d76mM7j0e/K9Lh3bHwl/4wpfG/LPDvP5M1OcYlPYa74qb2v0s+95SGh3f1dQqOqqOKT7Tj+G1114bcZzf+973RHyd4tJ4WnK5/8Z1N4z4uheHc+bNaxg9qZdffnmys/95m+GLYW/rmWeeqflc3LJzww03VH8+tf9qoikf+U+w3PDUy4u/1ex4ii6W28M6P4dF2BZfQ2zX7/3g+DLjP7s1oPv9e6vuYj7v7AqA+mhf4rJmX4Rfn/3sZ8f88xGa/0xPjPP39EwM4l6ZXl90Ojf27LNh4sWSN938Nd/T0hfnzzPPPHPE54Yamh9f0f7DD1f7CvW5557ryeOGGu7IUVxr/VmOr+e+994R/82q8f0mO2o7WvkZaBJ0OUE1nD59OuzY8YNk9+M7cHD9979f3VYRiLSs37B1xn/2xL5Nq6vWe7YpY/Zs2VJzP/bj99eEwqX4vLhQxfCiPftAXNVv/97vNvnP3tDQgWbtrbfeSm688W/jfj4dT+Hq4ONPPFE3g6zZV5GMATC5LMf3M33v7QML88wzz4z6fYvndLkIO9q77LFH3d6Ew0eRtdFmzZq17vwPv/DAmDcyeJ0Zf9ajqr+XtM1zDQQPt+J5N3ED4sEHH6zWwomWceedd45pRMQ26O2fNDREuBi+bdu2Ub8XgXniiSdWfR3pvXVfQ34OinHdfNo2UAnDh+U7hquqr61Vt0QCyCCWvF7upgOX+J0GBgYShyq0z+Z7d26uyZpK75WnnnrqiM9/8skn6/4Qx1NOPEHvyTBqdgz9iYux2Pns5z8/aoJwNm1Xfr/oEJI4h+LhOD09NfHSdyFWPz5/t+T32fTQWnY15DsZw6FDqQZOLTfccM2oLzg5dAhQmIGBgaQcY7XuSHJ9C4QBaKqfZ03tWupOKIwUGA+zHUdxDfF5nn/++bofXFf0//v93/KF8K2V39vR0dALmA0FzsNo34C0jUhfR5NzZ8PP/jTK09zqrn3gQXWvV+z2677N4mI8bdu2rX6KjhDeRphvU61/1NLLi+WFDx+u/f1vfetbY375fP/+/fXgAw9Our2oX3Jvly5d1P0FlOna+/c9PKsW4kMvvvBSQ20S0LMGz3C39PpWLwqv6Ohee8f2uq/lfOcffrihOlkzzt9TjLoHKn4ORn38GyVeHvL1r17wuAqBq4l0m5e1Wz379j0m7mo2NjY26vulR3G2bdvWsdgKlOJaQ9e2bt26qq9vfN0fsdhK53iMJvP1djAQ79syMjLS1Pkv8rquokXH8fDg1w3+91NPPdWQv50UNY1C1TzDxa+zP2fpNvXWvmyFU3znnXfG9JSsGJdeemsD4UaL+RjFf8b5drD33nuvIU/V7Xzw4MG6X2NEJ5fakNGKtR1Jxp9duHBBFM9JPOWTnkXjnMvt39/T8rqvKfW0ktArqfymvxhLy7G2P637Oj8/L+rulv8s7fKcW+ljLvw7d/ywafOv2kE5m6or60HeRLx2+K9bGLUTPTc1AzB9QieouhMvMBr1ORRCjKkIsSmGScfT3l49qihXRk9OtbdvH7GtrMceeyy5aNGXR3x+zQ2N4ua4NtM4HmoWwdcwttzx9EsvvTRi8R4XQzF8up5CJ+Msz67mGJj3r7/8SvLiW15bD3YqP2ZzHWFK0h6bf+TIkRGRfseOH/RHkPHtt3+9uiXw7//+7/VcOhHn7YYbrq95fyvSxDzMshGNr7/ttnqtZCv6Y/ZPxnDZZZdVn8+v1yxV+6iI7bN4zjnn1N2mo3bttdeO+uyjR48aau9Jv5cMOEWYk4o1V69p6v1dImxj2Afj2LFjI4J6T+Ac6uWf//u3a+7H+dsyXPKmvjE4v+4x48+C0P0ojjuzv/fT7Nj6l3cQ5ryP6zvQNBblf5PDI8J8e1oE7OyH/6vbC5p6vVwbIKuETvAf3uLTyHDzjjvuqLlf9oqIAUQ91YoqbTuVovf2lCWLFi0a8fnu7s7mL6gIotJXrXX1Q2ls1LZCR6VbM9544401/4a4d8L4HiIQuPDCC2vud7WvpWdaD8fTca4N2cyLYfK3bt0a9TkuZw2dWvryhW5t1A4n67vnXvde8vq7vludNhQaGvzP6vYQ0jYwXSHi3LslhmcjxlPOekXwFMHE2H/FAEwTq/W3z8sIIW3Xv2v/5mzdrl4Spk6deucVc64t8bTQ8XOi6Sg3XH/99R3vH7nxxhtr7h86NAhw7XhXxWw/fkdwqWTqYeI30udj9gzC8kSNbW1k7Y8hfj8xjjcE8f0mO2o7WvkZaBLh4jXG92uK6/FdfPHFY25vbAvToUMv8vIhOC4i0Wq80yjCcSO2+mGEr7mxht4sYLmr52B0+yPchx54rD1P3ot4vXb5s98t6mvOWrhGnB9p6RXClRZdflktwqHE4s8tvmwUNwbw+SMK9oUI/gLxk99XdMeNlhjFs+ugxXLVKlZXDzhWSdCIOL+J3ob4f2f+ebi3Kxq+2IXcTz755FG/Z9I5FCUOlc8EwLzFTO1Mdc40ep4zlfhSLQJBY4wdP96wN86fPz/i7+n3EYumNxC+5fNQ0/U6Wri8NN5Pf/rTiM+N+BuN4YjTK2oZHl3P0u3qvffea8g+9Tj+9re/NdTQ6NChQ2P+/Xq/cHtu3brV0P0YUZR48dfL4YcfXt3ekJ4fYWFblj0M1GhbnqWOHZJ13po9/sILWr5k0aJF8x6n8zJ6GtF5F4/gt12+xBhx3MHDhw91vG/olpwfbWeq2jGEAOYnttV7b6nMgb6+6lYCaLb4nBaA7DthcH30e3RpMYPHLipPek1P1NsdN+ubS2daF+JvbdGixYQHGT7vuHFfGTrV3b6B48dG/W7IMT00fLrt+vWrLfyS5NpLLo36vYjfdnSx3NnGPALzwvFr835dV5eXRfCUPg6yIh6NIM/dddddHejvt4O86dYvfWnEG3Y0xXPj448/Xvd5L1myZMTzs219AsALxG+K8NFlWJwY3G5qKeYZzJiztO+H/L6m9vTlf3lpSi3EuQ4dOtQ2+7rF9zwWhV6pP5uZDuA/ktqztN/zfP3jr/4TsyVz+vTp9qxPF8Kfefm8PI2fG+ecTr3mXfzmycnf/73f1f3+l4t46d6+fXu1I8+BAwfGfP7wQz1v+ueee+quyyXct2aM4lqfxcpvr/13k82Nwhzoy4vUoevAhfkXosPdr3/96yM+f2tvv/12w5eW2qLjaQecvvv21nxe+uZev/CO6hThgQMHGroOXxP10/U3Fd0IYf369aPev6H/jgX0J20D1TD7KsSPz/yZ44Ohl19++ajXk2xd96VrR3YfMfL1N+rZWLiyl7jwVPd6G37i5u89UQ4fPmyLtRitI+LYIoeKlyxZMrqFj+9ZxP3eRNZW+qtqz4j1dI0oTmMMwA7BlcQ9ikCd3fP9Lz4/5tvf1DKfTKZ5f+KuavTZ+vWr4a6xRyYdW3TGhyLzwvFfl76mWnR5vqrX2WBDHFDcvn171N9H1/pDMgf69vPfjjf+vfr3ie/96l9q7qeQ0M4jPq9du3bM53bo0CFbo0eAatnsrTD6z9AlC4f+5eWJ4R3e+MUv1j08XLJkSeP2k/fZfWfthne1TB8qFl0+/OD4qu96xz+qfk9L6BxNJ0RUwrB/nn322aoFS8Ze44lzz/y/rLzy1sRb2P09lM2fP3/U+7sEjo6Wp7ePivM2do3oA5WfHX2eYrj8+x/DMOecRry3b9/O1Nc8dOhQ2N52DbuHCa2rPf0I/L3EWqjdIBRDK+laUf9fjrWxcrL8OjPqqfD+53/+Z/W5L7744qj6xNfqAVSPr7+M8D7V4bSm5T9fzz7wYPKvf/+7mrUVk/bnnR3t/IKmEWye2OWzOILXqO7TFL59wS233FKd7wX0+7t07tzX6Gjf17fuxlQBSMOrpkWQRvNPHI+JmPKtF6PG33//hyPena1Q9fDhQx2P0RK3T6vPhdBdDv/oETz7+JOGQrbh5jM32q/+/in/N9X9HHrqqdUtfkn3bceXh3TR6flsV2vHf+TIkY4WXitXrqx7v6v+wOcQYuO///u/191W9uDXjVS8+rL/OLGfr2++tPhrV4tM2mNpS95Z3c7dQw+Nek2vt7jft2/fKdz48Vk8p1F5/oUN3SRv6L4J3n5d3V43lAO6ifn6Ef0aohN0J+yj78TgxHucxphvSdOUdZnT+8qR/8yf91B+fo2uB6AG416JhYmi3FYjVltcfBkyMuLDikjaF8PPhchxPPjAgyds4SPtmfTQ2KnOcYtLpfnPqbf6WoT7wAMP1L2vCeHpOkeToiz1fauvp6dR3drxjR89VJ3vIgoYnz17VtQ+nPj56aefrs73Wr58ec38kH73D954I2RZvH5dR+fQnQHv4fPZ+17VvkJ0dlNIuuc+97mG1syuqd1b4uF6bKkpc/4eXR/rGSESPtUTN4m1Xn311aN7M77fP/Hv7dTJk2vOo7i3aIeOwsQh9NB/Zp3Vjji9+h3VdGCfaXv1pQ4Y4qUjrp37aHzvW98ad1vJo+BjiTGWNsD4OQ3vfuSihXfV6MTRfD0dnU8PvP9+X2tYrS3CFoqf7P/ykY98NOrrR2dHdUX43VufMvwfT7hEl7Avg6cwuheyi+oiTK+P3LbJq0vPTuDvBcL91b6+tR4j4n64P80jSEf2YhPPe/eY6yaq+2gBj5ckzvW4zXL/ptdRns8TV9GfGEcRlhbc7ypV2odttBUU7Aueoqvgr7jqqqs6+zIBaop7s3RzuH3Xv3vB4yorXO7r/TmXXnrpqPd3L7jwK80P6uA8kLy7ZPl1RH/3iP8F8e0jtxDlrqEwuw1CGm5Lt7XvM0V4DUeOHGn6+VnO9+tx48q9Nl9f+95uvvnmquvQcNxwyJBpHU8xfNpwTVj79z40qlvNP9mvp71H6Xrf2NGHNWvX9hx/Dx8O9XVdrz7Xkqd6wiWI39ShhAeeJuxLnPMvG17l5/Rce+utt0Ykx3v+5VaY8dTfS1di5ObPfdHTl6/eQfbAU2zHa5SLFkI73FFfS8dhNJN1QYfR7hQ/13V0jvAM8p+90cLQz9G3b/9GzTaiCyM8vwcffHC1D7uWuLz2db+3tsz5kbb/YQyr1J8dtz1BqnukYXg2oRjvsQ3f7Gj+ie3f3wqbuoA79ms6iw+fww9PnTZCWgRPPX6O9bEBnS/efT5tgHYRykJNk6Z/bCYoYCs1y+kjR47UZC7I5ciVz0t0+QwfPmmpj6WmL3fRfb3zzjvV5w4PD0/sdwZ6xVeGxP4u3VeN9RRhAcS9/LoWL/zK+rrt27fXvW9do6cfv9/S7PEXX31Vcj96G1XoUsk8DjvC4PgKu6P9it+i0pvduZfV9fvG7/I6ykp2+mMZ+I/Hd88995z6aqPYBr19QZyy+ZmxTjDj7+Ieb8MNO7+vaf9+V78P3voPLmpYNXXGf9s3j5A5fa+0Nzo8Iq8Nr8PLcqLvrQSHc9Vh9Apm2m+ll8s4f87Pz0u3p1MUz3H4RqP6e3jvd1kegrgpOzYWp2aH3j4EPhf3mfbNrr7jyRd+Sc1rnUGWnBz6qY4ot3auuxS+4TGejrb9OtLTM6Yx/0JhdDulGfOW+GzjWlfR4S/eanDFca0E8tWWD+clqvCdfVjX+9YKq6dbfh+0D1R99s6dbv6GsxnyYSWAahnITk3z2+vrOPWToWxkUPWpM0rku+/ms/M3xW2rvsQ8zLKqZ+Gdd95xP8bqjTfeqLtOWfZ2vcTynFvv1Bbl7YV39qy0uz4cw8O6+3d0v23nIDmycjI/2tb/Lvqb0rHGz9HT2LmStlWQl7hWXjr6XdnG3z1hD/zbsWPHRvn7SZk/f37dxyLGK+u+1W3Jm3rG4Pwqqxv4hysmpMjLRv3e2mAeX9fQ7Smv+5ruNn0uE9O8W86P9H1i9EtvfS4dTK+44opOaPv5+fmus0dW/oODxpIUQ+Y9Tlt+Pv1Nzf22rf9cVUfTSy+91DzGRLQmvEj4FJfTOxr61+e1qnudbdu27YjPx7/+9a+N+dz3hUnxnkdNmTLliNcYnQBypI84Uf/eI7XvbQ2fY/26db8qC+DcxP0aUr5ut/rqq8f9m+Q9jnXdaxWxqya+1jUNd8QNfMysL4da9qM6oAd93XLVVTmP+CVupPRe9T4OF9vq/W0vvgymW38e0lXPbhNY+7rPz8/riot9X0tN9Llly5a6dx87dqzu4f57eHd3d+L3Qajm7/L7VxTL3Vu2bJn0vnn6y6+MiqKfPn16zOeSdnT4vR1P8ZyPDz74YEO8L3jiiSfGfK78Nzxz5syok6g5++yzM5ePrKM4f47deuV/dennf/az/RWzKbrlM9u0f/+orx89gYjvQWUbxPkec8wxt+mFC7/UHJf4ur4yzPvTF/M4M+r3zwhQx2CVfLM4Vxr1w3Ls2DGxN5r0HPB6vMfe0hMe8v1sfj69+OKLIz4v2/S0N9VnOtP8+0lDnNvJLrjggvrvJc659Fdex7qXXnrpqA/Hz3/+8zV7McK+iJnhu1inGNXfzeIYkxS/3tKlS0f8c9ddd414TAfBPfqZ77vvvvp5Wjdi2BrhJUa1uGJEt3bs0EMPHeU+26oA+c0PiM6uCb3BUyzgfVxN0HFLPkPx1l/K9V0U6zsO1vOslV/uws0jPv9VX19+fi2+3/7E+qnt23/f9O7z6g/4bx5ZGL0qPz+xZFdNDcb5t5bSOnrV1Jm3HhTvnzvUm/fv31+z/RMmxl3hJ873idHz2TPK66jWRx55ZET4Hfdb/N1uDkWvLz2fdFfDOA81f+6N76x24+zsk77+9yXSSp9P0wmWcsL/jZZGX8SHHnpoxPsTR9W0tHH7fSzVNkT4S4Jn9uc4vo7wjQ4nYQwX31oUR17z3xePPqrn47r50tHvS7dNO7f4vX+xPkVJ29HCwevY6Ry+Q4cOjbmutEXHUxjD6dOnG1rdrS3Srllsv9d5N1CqTz312Xj3K08ldc16hvu3b0z7q5W0A0EvnKYGOsy5vKJYfnJ8WW5HmHI4eejM3pP3/Xs7J2bffj4xr+nMmTMjPvfll1824jE6V9OBAwfGfTzi55Xkfxf3f9RQ4M2f+5LqvTn3vmP6ulccM5fiul9i/ly6jv1nnnkmmTVrljr/eaONNho9J3/+kyaPGWq4IfW+4Ydc68/JzjvvHPE5d4//bLr++uvb+7LM+V4P/eWXXx7xeDMh2nuwGm6qqT/XnA/r1q0b9Xw7fvwg2fV1f+YwgH6J+zWm9/XkK3/TpUNjOS/H5Y/Huc+GYnh4SM7Pj85Njq9/bOtHR8tRr+UnT57M8Yce6udrLl68uO7rTNsh4VjTVvCcOnWq7jv+4osvqr+/xx57rPr/aJj4sT95fv36Ue+vYfUr/P0pHuPaSy9NfL6x6t4hTuhqnvzsl6Zdzsp6ijGvVtJfVx5NdtrF67wMY8Wlbe2L2mi04ByOzz33nPp5GA7VjXNS4mpp5Svs+vquxDngU9sjvJ6o/eRl0vnzLwufh+pcmsVxuvhYz0Hojxc333xzzec6adPjL3al1EdBr3T2bAi/KdNRHTEx2Q2pTX2gl18yxdUWgrSKr39sxPebbM+YYWmCkGfs0UcfXd3P5s2bGy4mI7eL9IxeG0IPK5b2e/bsGfH8dEX4sUOHLN0e49smm5M4vrvOMW2P0f4ejktYK4n7Y4h7ozgt/qJ2LWg5rclWHdRd3jY4Vf9F0dXUn9X31AojzH1q+rZh/OXjIa4fMoTvj3cQ6x1G4y3u4dGjR8f0kP7a17yPPTfS0w4Lw59nNPKdjzzySEvuKsa3RNhrRCcnGV/NQmW+Nz2LQlnN9YRTfOgtXYfXXhp0Pp1Tmd6Hn8vZ2R3v69Jdbf7a17w7mn99q3scGn4cVC1+7qbi9+Z+GfX4D8V20tz/3nvvjXh9wz45PBR3bK/7u5Z2n0PZKCK+i2PLe9ddd43oB3vttdeMaPQVuM6jvJ6jvJ9l6Oe2dOnShgfeh/Pz8xP7+F/s65H4+Vk58fOSwKTRV+wF0zM5LNz8Nz0+wpDUv4fe8pY3V/u4zZgx4y1NPbP4ehbPddkJ+hK3+nWaTvj6lpPf+ObX6mYvnSf+uj+9RwP6+nb0N7PcekknhvQ6vja8+uqrI95jW7f/8cc+OuwcjPgbzWGAL0Jj9SP6M4g+178+9Pugr+8k7ecj6uz08L7DRJiXth+MbfVee+01I1Zteb1m2L+B0P66vcgnxY3HvA74n9V5Gcb6mu3W7L333tE8siiOje+999546vuydOnSmvo0Qp7VlfcfvV5HLVw5sTQ6fK3Zbj0DFGHcfuJ+ia+7v4eyWvoau0j5/Yp++/z8fNTzKV+L3DrgR3PPnuvmr+keRnR95R6vsw+fHwhHjhwZZaZrN2jfj7QXWtq2Y1pjy+MwrsHbnmYtHVGg8zCW4fPZ0nYk6dvPv4yI6tRb0a3f1P5K0n4+Yj3J9gVp9fXQz5/10PWVtr8e6s9zvi3CJ+5XeXQcR0v0M/7sl/fgwYOjToc5ivdSeiH0uV7/97f07HdqL+jrtfFct1xXt9fVNS7FIwj1nF3H90H6eYjvD7lz1KGfW0rT4ndqOjfjoOzhZ2DYePDb2PB6h6EtP/NfH11jsMrZDYvC4xUa6HyK8OHu7u66dx/b2ylM6M+G6wY5pd0QVm8nDg1we3fo+njb8Pqrvr20n6Si77Fx6NBxRz3eUydPjvncupZjHkEuhngfb+2QcL5B/8Xh1KVzKvoOlN7TFENrOzoa9sdHxfmaW7xG9dxTT6f7I4g8vupjZnlRfs7Hhls0/LlDh7u3/+fWv7+lj5uhyysNM1T5fPsvf3HcneDly5eP+twf/vCHo95fwwjD4qcQ13+uFd3xfpsfvj/3qVOnMjsM3GxEP6gk0KahvwfG4Y9ll91MfV+Vr29+fk7HYVxTL7zwQhZHN8VQ7qF9J23J8bXWjLiA49bPb9iwIc4Zh4cjjvsY3THgWZxLm/Tt3bv3qPffyy67LElL83X3Z5PeDl//2IhaT3GdK3hKi7F9Ruq+DH+0eTl37lyysWth7mNs2nSPvpp70P+OSjp5A9Jxvu6N+vWvH5UeV6T11ydQPNyJZZuNKR5zmifO6fQY0V8ZetWqxlo8dvbIupmKveQnM/PW0PsFWe9juJweR1Gh49e+JsL7WXr+e1n6PtA1JLS33qneQzjxsrqKsA369e8fcX0xgD/3WBT2R4q3V8f3W9h26esrcr99mz1+ko7DeO6vvfRSEkc/uq2AT9WvV50ztqv7elYHPx95vd6R69svf/mL12S0wQxNfUaNnqKz4D4d4Qv45hY3xsJ7GvJzPk7jO9V0PJ49e3YNh7J+/fpM//vcc8+tqefj4Vt6Hud18PgvfWnU+7N1X97KcO5x33ebwY0XlW+JjwefhlGjz9ns2bNjvl8nT54cEeu5td9rnDeL+zr6fI3sUmtHbwdO1e+xPF5n7YqP07h+uj7Pz8+P+trR6Rjn5ftdXOvT+cL9+nfd4kbfubqeb8+fPz8yWC7G98P7dQ1J+aQ+i/PpHSDWeel+ia+3xltvvdXw7it8fWtbhL94jcw99dRTI4LlHTt2NOTrR69r8f3u+2ecW8f0nM/uMdx0vfvuuzV9DhGms6J9jzzySGcL0YfPbaxQbO3fP6LjKN7LUfc6HUG4bfIltq/N8MF/dD0UF2Fj1XguYh2Vh7gIn8POcVyIXwQfKzpe4rrTRdB4OqvC58DeTGFPtL+H7xOdK+n3FudYPP/885n51k6GaU99L6ePv7uhv9sVrYLynIwHTqu+veptl3Hfxbifuc5t4sSJms9v1qyZvlfItvj6Ed3TqFl47dq1DXm9GI6yNt6N4yxJf1TSxXJ0u1Jn5qt38vT1jd1K0B12jla2hv7555+v+c0nn3yyYfQfV1jX7eu7r+8Vcf29Kx6r5N55552GfOfcccfd+pD3d9ZRp84fHxo6nfD5vO9DPG6o8/x3bni8Tg2f09OTWWiP5bnnnhv1+66t7Hs+bdrViFcM09Ep7s+Mqa6n+vqSdlj4nqX9q6JhOXny5M6XMt+tP7d//4uSvq4h0fWUazofR79L+3nK1/F40rfl6JGjNh7r6y4uzp/j9Aj9jXr08Ld/5Wmzzv+0f8Fv6uhzN9u35UGHPpz1dT1E1xP1v2v4n+I6mx4n+pr+R7+WOI9j6x59Taz/Qo7vmux6Y2ud6Qn6x02dOtWw+U/afjR5GXfHHXc0fE6hoV1GqR2erztM2A9h22fDUyvny8q3XNfR8RJeGV7P5MRb7DVK17Bv+3a2Ty5cuLCmz6C8nsdjyV9T03uavq7m5+c39D7Dcp7F85w+fa4lH1s0jrB4Ooaj03SiC9DrEe7dOhrIY4lbfj53+vTpNQ9zeF4LL7ywLkgd0feh7TVCku6I0wKqAOWQRQmAZJsPw3hXFNt9cS8QttWfWTRN1HH8R+KmGv43u4Joz/cYb65sX1uo8XQyXlxp0b4ghT6P5XGkRde2adOmUdffn7m6L1v12PZNXcfj8uXL4/7X7mJs34+4rWf0esA1ldZkWwVHjx5NfnD1VcM7IK7AdZVGOtq7Tz31VJ0fQvm6N4ri2MHobOS+P6LjJbyfsej1TQ1kP8P9mUjHaTSrX1YvAq8azHm4g2/hhb4bw2FuT3nZXlfs7Jf/XNX5r66vt4NeL9yW9PO0ftL2q81MW3qYn+Z7Pjk5abOx8mYO5/7o0aNjfpd6K3Ge9Mfl+I9r/7ztN+aC7CnOn6W8f6+f+L4yuM0mTpw46vN7eHfTGxr6S1h8nYjvR7JTp07V/FjTe9j+h42+J9/7DW95y5vq67r+8svHfL9Ln7eI2P66z/uWLPGu4pRTrLW8xjQXxrJhT89r5syZVc8n7fXiv4fbnEPnxJntCcJt4eFw6FC3nS7lXIszGvS+hrSM8vI7c+fNq/k3U8f3FWFvcljd4u8l3bdrf4/q/BusmKd9b6lP6s/HyvSKs3v06NENi/Ho0aN17/v06dPVM4BHev7lh6f4eqoXJr6fdMf3Pe7m3yiOPUPcl5kPbe9VfR3nxy8/dzF+4MABw/LR3xWhXd0wI/1nUj5pv2/fvobVjb63L7jgghqzFvfo6L+eGGJdX2v6XF0LjhfN7CeuZxjXG7JFDVWYn+M/tiX6mYjPZ3lz16gfi+uhvscnTpw4orpQG6c4fpL66R9Dl5fc2wcfuN90Do7qxv3pPYw4T0tL+66tvIa0PPLQQyP+eWqic6WtTc8Sj8P1S1cv6lA6i+vYQxt+l+gI/vTTT9c8n9WPt7iy/qz4y/4VnfV1l7Rfh72xpa/1f4y4p//4Yx+t+9/K9Lp6wfTPTv7ql7+M+LylL5H4eQmvuWOPPba6tVwe0PKNjdU2oNub6vDaxkLWY5xznDhxom3P0ZWmrwOH17G57u+HhtfX/ByvLzqaWq+9U2ea7wo8vN2rPKe6JuP3lPZ7VrQvUflUu8+x/Nprr5nW1Fnje3r27Bl1Hf1K0gdNzeL8e7juVD9guNnV9JmWPnaen5/vuFnd6PF66quvzrzHdNyefPLJmqwHf6f2d3HfhiXu20znn3vuuRob5qefftrWZClfp+O7crwFsb1H6l0eHh5uCP/AKPTz2A6Z7u7Sb7v11q+r21fSM1ic7Vk8ljh/YxfSzp07R1xf1uG5jePV+/rTf6a6Tbbqm1vt6+3g+3SO9dJLL9Xcr83slY7DO7Qh0nKvrwY6X8d77rlntAefLqL9e0f7Lr30UgZ5AbJgIix+gT0AAAAgJzoBAAAAkBOdAAAAAMiJTgAAAADkRCcAAAAAcqITAAAAADnRCQAAAICc6AQAAABATnQCAAAAICc6AQAAAJATnQAAAADIiU4AAAAA5EQnAAAAAHKiEwAAAAA50QkAAACAnOgEAAAAQE50AgAAACAnOgEAAACQE50AAAAAyIlOAAAAAOREJwAAAAByohMAAAAAOdEJAAAAgJzoBAAAAEBOdAIAAAAgJzoBAAAAkBOdAAAAAMiJTgAAAADkRCcAAAAAcqITAAAAADnRCQAAAICc6AQAAABATnQCAAAAICc6AQAAAJATnQAAAADIiU4AAAAA5EQnAAAAAHKiEwAAAAA50QkAAACAnOgEAAAAQE50AgAAACAnOgEAAACQE50AAAAAyIlOAAAAAOREJwAAAAByohMAAAAAOdEJAAAAgJzoBAAAAEBOdAIAAAAgJzoBAAAAkBOdAAAAAMiJTgAAAADkRCcAAAAAcqITAAAAADnRCQAAAICc6AQAAABATnQCAAAAICc6AQAAAJATnQAAAADIiU4AAAAA5EQnAAAAAHKiEwAAAAA50QkAAACAnOgEAAAAQE50AgAAACAnOgEAAACQE50AAAAAyIlOAAAAAOREJwAAAAByohMAAAAAOdEJAAAAgJzoBAAAAEBOdAIAAAAgJzoBAAAAkBOdAAAAAMiJTgAAAADkRCcAAAAAcqITAAAAADnRCQAAAICc6AQAAABATnQCAAAAICc6AQAAAJATnQAAAADIiU4AAAAA5EQnAAAAAHKiEwAAAAC5XwIMAPsqSMI3Lp24AAAAAElFTkSuQmCC";
const D1_SCHEMA_STATEMENTS = [
  "CREATE TABLE IF NOT EXISTS app_state (id INTEGER PRIMARY KEY, data TEXT NOT NULL, updated_at TEXT NOT NULL)",
  "CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, full_name TEXT NOT NULL, login TEXT NOT NULL UNIQUE, password_hash TEXT NOT NULL, password_plain TEXT, role TEXT NOT NULL, platform_0800_id TEXT, nuvidio_id TEXT, must_change_password INTEGER NOT NULL DEFAULT 1, is_active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL, updated_at TEXT NOT NULL)",
  "CREATE INDEX IF NOT EXISTS idx_users_login ON users(login)",
];
const INDEX_HTML = `<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PORTAL DE RESULTADOS</title>
  <link rel="icon" type="image/png" href="${FAVICON_DATA}">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="/styles.css">
</head>
<body>
  <div id="app"></div>
  <script>window.__brandLogo=${JSON.stringify(FAVICON_DATA)};</script>
  <script src="/script.js"></script>
</body>
</html>`;
const SECRET_KEY = "pulse-ops-local-secret";
const DEFAULT_PASSWORD = "Trocar@01";
const PRIMARY_MANAGER_HASH = "f18a137a143dc89817660f864bc973b0$6e3ebbd96a5a02981368b64d2b039dd93a52d901d44f73bc5f1d96797760aec7";
const STATIC_FILES = {
  "/": "index.html",
  "/index.html": "index.html",
  "/styles.css": "styles.css",
  "/script.js": "script.js",
  "/logos_KR-02.png": "logos_KR-02.png",
};

const isNode = typeof process !== "undefined" && !!process.versions?.node;
const isLocalNodeRuntime =
  isNode &&
  typeof process !== "undefined" &&
  Array.isArray(process.argv) &&
  typeof process.argv[1] === "string" &&
  typeof import.meta?.url === "string" &&
  import.meta.url.startsWith("file:");

let nodeModulesPromise;
let storageCache = null;

function decodeDataUriBase64(dataUri) {
  const [, base64 = ""] = String(dataUri || "").split(",", 2);
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

async function nodeModules() {
  if (!isLocalNodeRuntime) return null;
  if (!nodeModulesPromise) {
    nodeModulesPromise = Promise.all([
      import("node:fs/promises"),
      import("node:path"),
      import("node:http"),
      import("node:crypto"),
      import("node:url"),
      import("node:child_process"),
    ]).then(([fs, path, http, crypto, url, childProcess]) => ({
      fs,
      path,
      http,
      crypto,
      fileURLToPath: url.fileURLToPath,
      execFile: childProcess.execFile,
    }));
  }
  return nodeModulesPromise;
}

function nowIso() {
  return new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
}

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

function monthRef(dateValue) {
  return String(dateValue).slice(0, 7);
}

function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replaceAll(" ", "_")
    .replaceAll("-", "_")
    .replaceAll("/", "_")
    .replaceAll(".", "")
    .replaceAll("Ã£", "a")
    .replaceAll("Ã¡", "a")
    .replaceAll("Ã¢", "a")
    .replaceAll("Ã©", "e")
    .replaceAll("Ãª", "e")
    .replaceAll("Ã­", "i")
    .replaceAll("Ã³", "o")
    .replaceAll("Ã´", "o")
    .replaceAll("Ãµ", "o")
    .replaceAll("Ãº", "u")
    .replaceAll("Ã§", "c");
}

function toInt(value) {
  if (value === null || value === undefined || value === "") return 0;
  const text = String(value).trim().replace(",", ".");
  const parsed = Number(text);
  return Number.isFinite(parsed) ? Math.trunc(parsed) : 0;
}

function toFloat(value) {
  if (value === null || value === undefined || value === "") return 0;
  const text = String(value).trim().replace(",", ".");
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseDate(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return todayIso();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(raw)) {
    const [day, month, year] = raw.split("/");
    return `${year}-${month}-${day}`;
  }
  if (/^\d{2}-\d{2}-\d{4}$/.test(raw)) {
    const [day, month, year] = raw.split("-");
    return `${year}-${month}-${day}`;
  }
  throw new Error(`Data invalida: ${raw}`);
}

function isValidMonthRef(value) {
  return /^\d{4}-\d{2}$/.test(String(value || "").trim());
}

function normalizeComparable(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function serializeUser(user) {
  return {
    id: user.id,
    full_name: user.full_name,
    login: user.login,
    role: user.role,
    platform_0800_id: user.platform_0800_id,
    nuvidio_id: user.nuvidio_id,
    must_change_password: !!user.must_change_password,
    is_active: user.is_active,
    preferred_theme: user.preferred_theme,
    last_route: user.last_route,
  };
}

function normalizeRouteForRole(route, role = "operator") {
  const safeRoute = String(route || "").trim();
  const allowed = role === "manager"
    ? ["overview", "analysis", "history", "admin"]
    : ["overview", "analysis", "history"];
  return allowed.includes(safeRoute) ? safeRoute : "overview";
}

function normalizeUserRecord(user, roleFallback = "operator") {
  const role = user.role || roleFallback;
  return {
    ...user,
    role,
    must_change_password: Boolean(user.must_change_password),
    is_active: user.is_active !== false,
    preferred_theme: user.preferred_theme === "contrast" ? "contrast" : "dark",
    last_route: normalizeRouteForRole(user.last_route, role),
  };
}

function normalizeDbState(db) {
  db.users = (db.users || []).map((user, index) => normalizeUserRecord(user, index === 0 ? "manager" : "operator"));
  if (!db.users.some((user) => String(user.login).trim().toLowerCase() === "joao.fonseca")) {
    const nextUserId = Math.max(0, ...db.users.map((user) => Number(user.id) || 0)) + 1;
    db.users.push({
      id: nextUserId,
      full_name: "JoÃ£o Fonseca",
      login: "joao.fonseca",
      password_hash: PRIMARY_MANAGER_HASH,
      password_plain: "Krsa@2026",
      role: "manager",
      platform_0800_id: "",
      nuvidio_id: "",
      must_change_password: false,
      is_active: true,
      preferred_theme: "dark",
      last_route: "overview",
      created_at: nowIso(),
      updated_at: nowIso(),
    });
  }
  db.dailyMetrics = (db.dailyMetrics || []).map((metric) => {
    const normalizedMetric = {
      production_0800: 0,
      production_nuvidio: 0,
      ...metric,
    };
    if (normalizedMetric.production_0800 || normalizedMetric.production_nuvidio) {
      normalizedMetric.production = toInt(normalizedMetric.production_0800) + toInt(normalizedMetric.production_nuvidio);
    } else {
      normalizedMetric.production = toInt(normalizedMetric.production);
    }
    return normalizedMetric;
  });
  db.qualityScores = db.qualityScores || [];
  db.importLogs = db.importLogs || [];
  const nextUserCounter = Math.max(1, ...db.users.map((user) => Number(user.id) || 0)) + 1;
  db.counters = db.counters || { users: nextUserCounter, dailyMetrics: 1, qualityScores: 1, importLogs: 1 };
  db.counters.users = Math.max(db.counters.users || 1, nextUserCounter);
  return db;
}

function normalizeD1UserRow(row) {
  return normalizeUserRecord({
    id: Number(row.id),
    full_name: String(row.full_name || "").trim(),
    login: String(row.login || "").trim(),
    password_hash: String(row.password_hash || "").trim(),
    password_plain: String(row.password_plain || "").trim(),
    role: String(row.role || "operator").trim(),
    platform_0800_id: String(row.platform_0800_id || "").trim(),
    nuvidio_id: String(row.nuvidio_id || "").trim(),
    must_change_password: Number(row.must_change_password) === 1,
    is_active: Number(row.is_active) === 1,
    preferred_theme: String(row.preferred_theme || "dark").trim(),
    last_route: String(row.last_route || "overview").trim(),
    created_at: row.created_at || nowIso(),
    updated_at: row.updated_at || nowIso(),
  });
}

async function ensureDefaultPasswords(db) {
  let changed = false;
  const admin = db.users.find((user) => String(user.login).trim().toLowerCase() === "admin");
  if (admin && !String(admin.password_hash || "").trim()) {
    admin.password_hash = await hashPassword("admin123");
    admin.password_plain = "admin123";
    admin.updated_at = nowIso();
    changed = true;
  }
  const joao = db.users.find((user) => String(user.login).trim().toLowerCase() === "joao.fonseca");
  if (joao && !String(joao.password_hash || "").trim()) {
    joao.password_hash = PRIMARY_MANAGER_HASH;
    joao.password_plain = "Krsa@2026";
    joao.updated_at = nowIso();
    changed = true;
  }
  return changed;
}

async function loadUsersFromD1(connection) {
  const result = await connection.prepare(`
    SELECT
      id,
      full_name,
      login,
      password_hash,
      password_plain,
      role,
      platform_0800_id,
      nuvidio_id,
      must_change_password,
      is_active,
      preferred_theme,
      last_route,
      created_at,
      updated_at
    FROM users
    ORDER BY id
  `).all();
  return (result?.results || []).map(normalizeD1UserRow);
}

async function persistUsersToD1(connection, users) {
  for (const user of users) {
    await connection.prepare(`
      INSERT INTO users (
        id,
        full_name,
        login,
        password_hash,
        password_plain,
        role,
        platform_0800_id,
        nuvidio_id,
        must_change_password,
        is_active,
        preferred_theme,
        last_route,
        created_at,
        updated_at
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(id) DO UPDATE SET
        full_name = excluded.full_name,
        login = excluded.login,
        password_hash = excluded.password_hash,
        password_plain = excluded.password_plain,
        role = excluded.role,
        platform_0800_id = excluded.platform_0800_id,
        nuvidio_id = excluded.nuvidio_id,
        must_change_password = excluded.must_change_password,
        is_active = excluded.is_active,
        preferred_theme = excluded.preferred_theme,
        last_route = excluded.last_route,
        created_at = excluded.created_at,
        updated_at = excluded.updated_at
    `).bind(
      Number(user.id),
      String(user.full_name || "").trim(),
      String(user.login || "").trim(),
      String(user.password_hash || "").trim(),
      String(user.password_plain || "").trim(),
      String(user.role || "operator").trim(),
      String(user.platform_0800_id || "").trim(),
      String(user.nuvidio_id || "").trim(),
      user.must_change_password ? 1 : 0,
      user.is_active ? 1 : 0,
      user.preferred_theme === "contrast" ? "contrast" : "dark",
      normalizeRouteForRole(user.last_route, user.role),
      user.created_at || nowIso(),
      user.updated_at || nowIso(),
    ).run();
  }

  if (users.length) {
    const placeholders = users.map(() => "?").join(", ");
    await connection.prepare(`DELETE FROM users WHERE id NOT IN (${placeholders})`)
      .bind(...users.map((user) => Number(user.id)))
      .run();
    return;
  }

  await connection.prepare("DELETE FROM users").run();
}

async function ensureD1Schema(connection) {
  for (const statement of D1_SCHEMA_STATEMENTS) {
    await connection.prepare(statement).run();
  }
  const tableInfo = await connection.prepare("PRAGMA table_info(users)").all();
  const columns = new Set((tableInfo?.results || []).map((row) => String(row.name)));
  if (!columns.has("platform_0800_id")) {
    await connection.exec("ALTER TABLE users ADD COLUMN platform_0800_id TEXT");
  }
  if (!columns.has("nuvidio_id")) {
    await connection.exec("ALTER TABLE users ADD COLUMN nuvidio_id TEXT");
  }
  if (!columns.has("must_change_password")) {
    await connection.exec("ALTER TABLE users ADD COLUMN must_change_password INTEGER NOT NULL DEFAULT 1");
  }
  if (!columns.has("password_plain")) {
    await connection.exec("ALTER TABLE users ADD COLUMN password_plain TEXT");
  }
  if (!columns.has("is_active")) {
    await connection.exec("ALTER TABLE users ADD COLUMN is_active INTEGER NOT NULL DEFAULT 1");
  }
  if (!columns.has("preferred_theme")) {
    await connection.exec("ALTER TABLE users ADD COLUMN preferred_theme TEXT NOT NULL DEFAULT 'dark'");
  }
  if (!columns.has("last_route")) {
    await connection.exec("ALTER TABLE users ADD COLUMN last_route TEXT NOT NULL DEFAULT 'overview'");
  }
}

function calculateEffectiveness(metric) {
  const actionable =
    toInt(metric.calls_0800_approved) +
    toInt(metric.calls_0800_rejected) +
    toInt(metric.calls_0800_pending) +
    toInt(metric.calls_nuvidio_approved) +
    toInt(metric.calls_nuvidio_rejected);
  const total = actionable + toInt(metric.calls_0800_no_action) + toInt(metric.calls_nuvidio_no_action);
  if (!total) return 0;
  return Number(((actionable / total) * 100).toFixed(2));
}

async function hashPassword(password) {
  const salt = randomHex(16);
  const hash = await digestHex(`${salt}:${password}`);
  return `${salt}$${hash}`;
}

async function verifyPassword(password, storedHash) {
  const [salt, hash] = String(storedHash || "").split("$");
  if (!salt || !hash) return false;
  const candidate = await digestHex(`${salt}:${password}`);
  return timingSafeEqual(hash, candidate);
}

async function digestHex(input) {
  const data = new TextEncoder().encode(input);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function randomHex(length) {
  const bytes = new Uint8Array(length);
  crypto.getRandomValues(bytes);
  return [...bytes].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function timingSafeEqual(a, b) {
  if (a.length !== b.length) return false;
  let result = 0;
  for (let i = 0; i < a.length; i += 1) {
    result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return result === 0;
}

async function signSession(payload) {
  const raw = JSON.stringify(payload);
  const signature = await digestHex(`${SECRET_KEY}:${raw}`);
  return `${btoa(raw)}.${signature}`;
}

function resolveSecretKey(env) {
  return String(env?.SECRET_KEY || SECRET_KEY);
}

async function signSessionWithEnv(payload, env) {
  const raw = JSON.stringify(payload);
  const signature = await digestHex(`${resolveSecretKey(env)}:${raw}`);
  return `${btoa(raw)}.${signature}`;
}

async function verifySession(token, env) {
  if (!token || !token.includes(".")) return null;
  const [rawBase64, signature] = token.split(".");
  const raw = atob(rawBase64);
  const expected = await digestHex(`${resolveSecretKey(env)}:${raw}`);
  if (!timingSafeEqual(signature, expected)) return null;
  const payload = JSON.parse(raw);
  if (!payload.expires_at || new Date(payload.expires_at) < new Date()) return null;
  return payload;
}

function seedState() {
  return {
    counters: {
      users: 2,
      dailyMetrics: 1,
      qualityScores: 1,
      importLogs: 1,
    },
    users: [
      {
        id: 1,
        full_name: "Administrador",
        login: "admin",
        password_hash: null,
        password_plain: "admin123",
        role: "manager",
        platform_0800_id: "GESTOR-001",
        nuvidio_id: "NUVIDIO-001",
        must_change_password: false,
        is_active: true,
        preferred_theme: "dark",
        last_route: "overview",
        created_at: nowIso(),
        updated_at: nowIso(),
      },
    ],
    dailyMetrics: [],
    qualityScores: [],
    importLogs: [],
  };
}

async function ensureStorage(env = {}) {
  if (env?.DB) {
    await ensureD1Schema(env.DB);
    const row = await env.DB.prepare("SELECT data FROM app_state WHERE id = ?").bind(D1_STATE_ID).first();
    const baseState = row?.data ? normalizeDbState(JSON.parse(row.data)) : normalizeDbState(seedState());
    const loadedUsers = await loadUsersFromD1(env.DB);
    if (loadedUsers.length) {
      baseState.users = loadedUsers;
    } else if (!row?.data) {
      baseState.users[0].password_hash = await hashPassword("admin123");
    }
    const db = normalizeDbState(baseState);
    const repairedPasswords = await ensureDefaultPasswords(db);
    const shouldPersist =
      !row?.data ||
      !loadedUsers.length ||
      repairedPasswords ||
      db.users.length !== loadedUsers.length;
    if (shouldPersist) {
      await persistStorage(db, env);
    }
    storageCache = db;
    return db;
  }
  if (storageCache) return storageCache;
  if (isLocalNodeRuntime) {
    const { fs, path } = await nodeModules();
    await fs.mkdir(path.dirname(DB_PATH), { recursive: true });
    try {
      const raw = await fs.readFile(DB_PATH, "utf8");
      storageCache = normalizeDbState(JSON.parse(raw));
    } catch {
      storageCache = normalizeDbState(seedState());
      storageCache.users[0].password_hash = await hashPassword("admin123");
      await persistStorage(storageCache, env);
    }
  } else {
    storageCache = normalizeDbState(seedState());
    storageCache.users[0].password_hash = await hashPassword("admin123");
  }
  return storageCache;
}

async function persistStorage(db, env = {}) {
  if (env?.DB) {
    await ensureD1Schema(env.DB);
    await persistUsersToD1(env.DB, db.users || []);
    await env.DB.prepare(`
      INSERT INTO app_state (id, data, updated_at)
      VALUES (?, ?, ?)
      ON CONFLICT(id) DO UPDATE SET
        data = excluded.data,
        updated_at = excluded.updated_at
    `).bind(D1_STATE_ID, JSON.stringify(db), nowIso()).run();
    storageCache = db;
    return;
  }
  if (!isLocalNodeRuntime || !db) {
    storageCache = db;
    return;
  }
  storageCache = db;
  const { fs } = await nodeModules();
  await fs.writeFile(DB_PATH, JSON.stringify(db, null, 2), "utf8");
}

function nextId(db, key) {
  const id = db.counters[key];
  db.counters[key] += 1;
  return id;
}

function jsonResponse(payload, status = 200, headers = {}) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
      ...headers,
    },
  });
}

async function getCurrentUser(request, db, env) {
  const cookieHeader = request.headers.get("cookie") || "";
  const token = parseCookies(cookieHeader)[SESSION_NAME];
  if (!token) return null;
  const session = await verifySession(token, env);
  if (!session) return null;
  return db.users.find((user) => user.id === session.user_id && user.is_active) || null;
}

function parseCookies(cookieHeader) {
  return cookieHeader
    .split(";")
    .map((item) => item.trim())
    .filter(Boolean)
    .reduce((acc, item) => {
      const index = item.indexOf("=");
      if (index === -1) return acc;
      acc[item.slice(0, index)] = decodeURIComponent(item.slice(index + 1));
      return acc;
    }, {});
}

async function requireAuth(request, db, env) {
  const user = await getCurrentUser(request, db, env);
  if (!user) {
    return { error: jsonResponse({ error: "Nao autenticado" }, 401) };
  }
  return { user };
}

async function requireManager(request, db, env) {
  const auth = await requireAuth(request, db, env);
  if (auth.error) return auth;
  if (auth.user.role !== "manager") {
    return { error: jsonResponse({ error: "Acesso negado" }, 403) };
  }
  return auth;
}

function matchUser(users, row) {
  const identifiers = {
    name: String(row.nome || row.operador || "").trim().toLowerCase(),
    login: String(row.login || row.usuario || "").trim().toLowerCase(),
    id0800: String(row.id_0800 || row.usuario_0800 || row.plataforma_0800 || "").trim().toLowerCase(),
    idNuvidio: String(row.id_nuvidio || row.usuario_nuvidio || row.plataforma_nuvidio || "").trim().toLowerCase(),
  };
  return users.find((user) => {
    return (
      (identifiers.id0800 && String(user.platform_0800_id || "").trim().toLowerCase() === identifiers.id0800) ||
      (identifiers.idNuvidio && String(user.nuvidio_id || "").trim().toLowerCase() === identifiers.idNuvidio) ||
      (identifiers.login && String(user.login || "").trim().toLowerCase() === identifiers.login) ||
      (identifiers.name && String(user.full_name || "").trim().toLowerCase() === identifiers.name)
    );
  });
}

function matchUserByName(users, name) {
  const target = normalizeComparable(name);
  if (!target) return null;
  return users.find((user) => normalizeComparable(user.full_name) === target) || null;
}

function parseCsv(text) {
  const rows = [];
  let current = "";
  let row = [];
  let inQuotes = false;
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];
    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      row.push(current);
      current = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(current);
      if (row.some((cell) => cell.length)) rows.push(row);
      row = [];
      current = "";
    } else {
      current += char;
    }
  }
  if (current.length || row.length) {
    row.push(current);
    if (row.some((cell) => cell.length)) rows.push(row);
  }
  if (!rows.length) return [];
  const headers = rows[0].map((header) => normalizeHeader(header));
  return rows.slice(1).map((cells) => {
    const entry = {};
    headers.forEach((header, index) => {
      entry[header] = cells[index] ?? "";
    });
    return entry;
  });
}

function hasFields(row, fields) {
  return fields.every((field) => field in row);
}

function detectImportSchema(row, fileName = "") {
  const lowerName = String(fileName || "").toLowerCase();
  const normalizedFields = [
    "data",
    "producao",
    "0800_aprovado",
    "0800_cancelada",
    "0800_pendenciada",
    "0800_sem_acao",
    "nuvidio_aprovada",
    "nuvidio_reprovada",
    "nuvidio_sem_acao",
  ];
  if (hasFields(row, normalizedFields)) return "normalized";
  if (hasFields(row, ["motivo", "numero_do_protocolo", "data_abertura_ocorrencia", "usuario_de_abertura_da_ocorrencia"])) return "0800";
  if (hasFields(row, ["protocolo", "data_abreviada", "email_do_atendente", "tag"])) return "nuvidio";
  if (lowerName.includes("0800")) return "0800";
  if (lowerName.includes("nuvidio")) return "nuvidio";
  return "unknown";
}

function normalize0800Reason(value) {
  return normalizeComparable(value).toUpperCase();
}

function normalizeNuvidioTag(value) {
  return normalizeComparable(value);
}

async function runPythonScript(args) {
  const modules = await nodeModules();
  const pythonPath = "C:\\Users\\joao.fonseca.KRCONSULTORIA\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe";
  return await new Promise((resolve, reject) => {
    modules.execFile(pythonPath, args, { cwd: process.cwd(), encoding: "utf8", maxBuffer: 10 * 1024 * 1024 }, (error, stdout, stderr) => {
      if (error) {
        reject(new Error(stderr?.trim() || error.message));
        return;
      }
      resolve(stdout);
    });
  });
}

async function generateQualityTemplate(users) {
  const { fs, path } = await nodeModules();
  const tempDir = path.join(process.cwd(), "data", "tmp");
  await fs.mkdir(tempDir, { recursive: true });
  const stamp = Date.now();
  const jsonPath = path.join(tempDir, `quality-users-${stamp}.json`);
  const outputPath = path.join(tempDir, `modelo-monitoria-${stamp}.xlsx`);
  const scriptPath = path.join(process.cwd(), "tools", "quality_xlsx.py");
  await fs.writeFile(jsonPath, JSON.stringify(users.map((user) => user.full_name), null, 2), "utf8");
  try {
    await runPythonScript([scriptPath, "export", jsonPath, outputPath]);
    return await fs.readFile(outputPath);
  } finally {
    await fs.unlink(jsonPath).catch(() => {});
    await fs.unlink(outputPath).catch(() => {});
  }
}

async function parseQualityWorkbook(file) {
  const { fs, path } = await nodeModules();
  const tempDir = path.join(process.cwd(), "data", "tmp");
  await fs.mkdir(tempDir, { recursive: true });
  const stamp = Date.now();
  const inputPath = path.join(tempDir, `quality-upload-${stamp}.xlsx`);
  const scriptPath = path.join(process.cwd(), "tools", "quality_xlsx.py");
  const buffer = Buffer.from(await file.arrayBuffer());
  await fs.writeFile(inputPath, buffer);
  try {
    const stdout = await runPythonScript([scriptPath, "import", inputPath]);
    return JSON.parse(stdout);
  } finally {
    await fs.unlink(inputPath).catch(() => {});
  }
}

function upsertDailyMetric(db, userId, metricDate, values, sourceName) {
  let metric = db.dailyMetrics.find((item) => item.user_id === userId && item.metric_date === metricDate);
  if (!metric) {
    metric = {
      id: nextId(db, "dailyMetrics"),
      user_id: userId,
      metric_date: metricDate,
      production: 0,
      production_0800: 0,
      production_nuvidio: 0,
      calls_0800_approved: 0,
      calls_0800_rejected: 0,
      calls_0800_pending: 0,
      calls_0800_no_action: 0,
      calls_nuvidio_approved: 0,
      calls_nuvidio_rejected: 0,
      calls_nuvidio_no_action: 0,
      import_source: sourceName,
      created_at: nowIso(),
      updated_at: nowIso(),
    };
    db.dailyMetrics.push(metric);
  }
  Object.assign(metric, values, {
    import_source: sourceName ?? metric.import_source,
    updated_at: nowIso(),
  });
  if ("production_0800" in values || "production_nuvidio" in values) {
    metric.production = toInt(metric.production_0800) + toInt(metric.production_nuvidio);
  } else if ("production" in values) {
    metric.production = toInt(values.production);
  }
  return metric;
}

function importNormalizedRows(db, users, rows, sourceName) {
  const required = [
    "data",
    "producao",
    "0800_aprovado",
    "0800_cancelada",
    "0800_pendenciada",
    "0800_sem_acao",
    "nuvidio_aprovada",
    "nuvidio_reprovada",
    "nuvidio_sem_acao",
  ];
  const first = rows[0] || {};
  const missing = required.filter((field) => !(field in first));
  if (missing.length) {
    throw new Error(`Colunas obrigatorias ausentes: ${missing.join(", ")}`);
  }
  let processed = 0;
  let rejected = 0;
  const errors = [];
  rows.forEach((row, index) => {
    try {
      const user = matchUser(users, row);
      if (!user) throw new Error("Operador nao encontrado para os identificadores informados.");
      upsertDailyMetric(
        db,
        user.id,
        parseDate(row.data),
        {
          production: toInt(row.producao),
          calls_0800_approved: toInt(row["0800_aprovado"]),
          calls_0800_rejected: toInt(row["0800_cancelada"]),
          calls_0800_pending: toInt(row["0800_pendenciada"]),
          calls_0800_no_action: toInt(row["0800_sem_acao"]),
          calls_nuvidio_approved: toInt(row["nuvidio_aprovada"]),
          calls_nuvidio_rejected: toInt(row["nuvidio_reprovada"]),
          calls_nuvidio_no_action: toInt(row["nuvidio_sem_acao"]),
        },
        sourceName,
      );
      processed += 1;
    } catch (error) {
      rejected += 1;
      errors.push({ row: index + 2, error: error.message });
    }
  });
  return { processed, rejected, errors };
}

function import0800Rows(db, users, rows, sourceName) {
  const aggregates = new Map();
  const errors = [];
  rows.forEach((row, index) => {
    try {
      const user = matchUser(users, {
        nome: row.usuario_de_abertura_da_ocorrencia,
        usuario: row.usuario_de_abertura_da_ocorrencia,
        plataforma_0800: row.usuario_de_abertura_da_ocorrencia,
      });
      if (!user) throw new Error("Operador do 0800 nao encontrado no cadastro.");
      const metricDate = parseDate(row.data_abertura_ocorrencia);
      const key = `${user.id}::${metricDate}`;
      const current = aggregates.get(key) || {
        userId: user.id,
        metricDate,
        production_0800: 0,
        calls_0800_approved: 0,
        calls_0800_rejected: 0,
        calls_0800_pending: 0,
        calls_0800_no_action: 0,
      };
      current.production_0800 += 1;
      const reason = normalize0800Reason(row.motivo);
      if (reason === "APROVADA" || reason === "APROVADO") current.calls_0800_approved += 1;
      else if (reason === "CANCELADA" || reason === "REPROVADA") current.calls_0800_rejected += 1;
      else if (reason === "PENDENCIADA" || reason === "PENDENCIADO") current.calls_0800_pending += 1;
      else current.calls_0800_no_action += 1;
      aggregates.set(key, current);
    } catch (error) {
      errors.push({ row: index + 2, error: error.message });
    }
  });

  for (const aggregate of aggregates.values()) {
    upsertDailyMetric(
      db,
      aggregate.userId,
      aggregate.metricDate,
      {
        production_0800: aggregate.production_0800,
        calls_0800_approved: aggregate.calls_0800_approved,
        calls_0800_rejected: aggregate.calls_0800_rejected,
        calls_0800_pending: aggregate.calls_0800_pending,
        calls_0800_no_action: aggregate.calls_0800_no_action,
      },
      sourceName,
    );
  }

  return { processed: aggregates.size, rejected: errors.length, errors };
}

function importNuvidioRows(db, users, rows, sourceName) {
  const aggregates = new Map();
  const errors = [];
  rows.forEach((row, index) => {
    try {
      const user = matchUser(users, {
        login: row.email_do_atendente,
        usuario: row.email_do_atendente,
        plataforma_nuvidio: row.email_do_atendente,
        id_nuvidio: row.email_do_atendente,
      });
      if (!user) throw new Error("Operador da Nuvidio nao encontrado no cadastro.");
      const metricDate = parseDate(row.data_abreviada);
      const key = `${user.id}::${metricDate}`;
      const current = aggregates.get(key) || {
        userId: user.id,
        metricDate,
        production_nuvidio: 0,
        calls_nuvidio_approved: 0,
        calls_nuvidio_rejected: 0,
        calls_nuvidio_no_action: 0,
      };
      current.production_nuvidio += 1;
      const tag = normalizeNuvidioTag(row.tag);
      if (tag === "aprovada" || tag === "aprovado") current.calls_nuvidio_approved += 1;
      else if (tag === "reprovada" || tag === "reprovado") current.calls_nuvidio_rejected += 1;
      else current.calls_nuvidio_no_action += 1;
      aggregates.set(key, current);
    } catch (error) {
      errors.push({ row: index + 2, error: error.message });
    }
  });

  for (const aggregate of aggregates.values()) {
    upsertDailyMetric(
      db,
      aggregate.userId,
      aggregate.metricDate,
      {
        production_nuvidio: aggregate.production_nuvidio,
        calls_nuvidio_approved: aggregate.calls_nuvidio_approved,
        calls_nuvidio_rejected: aggregate.calls_nuvidio_rejected,
        calls_nuvidio_no_action: aggregate.calls_nuvidio_no_action,
      },
      sourceName,
    );
  }

  return { processed: aggregates.size, rejected: errors.length, errors };
}

function registerImportLog(db, userId, sourceName, processed, rejected) {
  db.importLogs.push({
    id: nextId(db, "importLogs"),
    imported_by: userId,
    source_name: sourceName,
    imported_at: nowIso(),
    rows_processed: processed,
    rows_rejected: rejected,
  });
}

function buildOverview(db, user, url) {
  const dateValue = url.searchParams.get("date") || todayIso();
  const start = url.searchParams.get("start") || shiftDate(-29);
  const end = url.searchParams.get("end") || todayIso();
  const scopedMetrics = db.dailyMetrics.filter((metric) => user.role === "manager" || metric.user_id === user.id);
  const todayRows = scopedMetrics.filter((metric) => metric.metric_date === dateValue);
  const monthRows = db.qualityScores.filter((score) => score.reference_month === monthRef(dateValue) && (user.role === "manager" || score.user_id === user.id));
  const trendRows = scopedMetrics
    .filter((metric) => metric.metric_date >= start && metric.metric_date <= end)
    .sort((a, b) => a.metric_date.localeCompare(b.metric_date));
  const cards = {
    production: todayRows.reduce((sum, row) => sum + toInt(row.production), 0),
    effectiveness: average(todayRows.map(calculateEffectiveness)),
    quality: average(monthRows.map((item) => toFloat(item.score))),
  };
  const trendMap = new Map();
  for (const row of trendRows) {
    const current = trendMap.get(row.metric_date) || { date: row.metric_date, production: 0, effectivenessParts: [] };
    current.production += toInt(row.production);
    current.effectivenessParts.push(calculateEffectiveness(row));
    trendMap.set(row.metric_date, current);
  }
  const trend = [...trendMap.values()].map((item) => ({
    date: item.date,
    production: item.production,
    effectiveness: average(item.effectivenessParts),
  }));
  const operators = todayRows.map((row) => ({
    name: db.users.find((entry) => entry.id === row.user_id)?.full_name || "Operador",
    production: row.production,
    effectiveness: calculateEffectiveness(row),
  }));
  return { cards, trend, operators };
}

function buildAnalysis(db, user, url) {
  const start = url.searchParams.get("start") || shiftDate(-29);
  const end = url.searchParams.get("end") || todayIso();
  const users = db.users.filter((entry) => entry.is_active && (user.role === "manager" || entry.id === user.id));
  const ranking = users.map((entry) => {
    const metrics = db.dailyMetrics.filter((row) => row.user_id === entry.id && row.metric_date >= start && row.metric_date <= end);
    const quality = db.qualityScores.find((score) => score.user_id === entry.id && score.reference_month === monthRef(end));
    return {
      user_id: entry.id,
      name: entry.full_name,
      avg_production: metrics.length ? Number((metrics.reduce((sum, row) => sum + toInt(row.production), 0) / metrics.length).toFixed(2)) : 0,
      total_production: metrics.reduce((sum, row) => sum + toInt(row.production), 0),
      effectiveness: average(metrics.map(calculateEffectiveness)),
      quality: toFloat(quality?.score),
      active_days: metrics.length,
    };
  }).sort((a, b) => b.total_production - a.total_production || a.name.localeCompare(b.name));
  return { ranking, period: { start, end } };
}

function buildHistory(db, user, url) {
  const start = url.searchParams.get("start") || shiftDate(-29);
  const end = url.searchParams.get("end") || todayIso();
  const requestedUserId = Number(url.searchParams.get("user_id") || user.id);
  const targetUserId = user.role === "manager" ? requestedUserId : user.id;
  const history = db.dailyMetrics
    .filter((row) => row.user_id === targetUserId && row.metric_date >= start && row.metric_date <= end)
    .sort((a, b) => b.metric_date.localeCompare(a.metric_date))
    .map((row) => ({ ...row, effectiveness: calculateEffectiveness(row) }));
  const quality = db.qualityScores
    .filter((item) => item.user_id === targetUserId)
    .sort((a, b) => b.reference_month.localeCompare(a.reference_month));
  return { history, quality };
}

function average(values) {
  if (!values.length) return 0;
  return Number((values.reduce((sum, value) => sum + Number(value || 0), 0) / values.length).toFixed(2));
}

function shiftDate(days) {
  const date = new Date();
  date.setDate(date.getDate() + days);
  return date.toISOString().slice(0, 10);
}

function cloudQualityUnsupported() {
  return jsonResponse({
    error: "No Cloudflare, a monitoria em XLSX ainda precisa ser processada localmente. Use CSV ou integre uma rotina externa antes do deploy final.",
  }, 501);
}

async function handleApi(request, url, db, env = {}) {
  if (url.pathname === "/api/auth/me" && request.method === "GET") {
    const user = await getCurrentUser(request, db, env);
    return jsonResponse({ user: user ? serializeUser(user) : null });
  }

  if (url.pathname === "/api/auth/login" && request.method === "POST") {
    const payload = await request.json();
    const user = db.users.find((entry) => entry.login === String(payload.login || "").trim() && entry.is_active);
    const inputPassword = String(payload.password || "");
    const hashMatches = user ? await verifyPassword(inputPassword, user.password_hash) : false;
    const plainMatches = user ? String(user.password_plain || "") === inputPassword : false;
    if (!user || (!hashMatches && !plainMatches)) {
      return jsonResponse({ error: "Credenciais invalidas" }, 401);
    }
    if (plainMatches && !hashMatches) {
      user.password_hash = await hashPassword(inputPassword);
      user.updated_at = nowIso();
      await persistStorage(db, env);
    }
    const token = await signSessionWithEnv({
      user_id: user.id,
      expires_at: new Date(Date.now() + 7 * 86400000).toISOString(),
    }, env);
    return jsonResponse(
      { user: serializeUser(user) },
      200,
      { "set-cookie": `${SESSION_NAME}=${token}; Path=/; HttpOnly; SameSite=Lax` },
    );
  }

  if (url.pathname === "/api/auth/logout" && request.method === "POST") {
    return jsonResponse({ ok: true }, 200, {
      "set-cookie": `${SESSION_NAME}=; Path=/; Expires=Thu, 01 Jan 1970 00:00:00 GMT; HttpOnly; SameSite=Lax`,
    });
  }

  if (url.pathname === "/api/auth/preferences" && request.method === "PATCH") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    const payload = await request.json();
    if (payload.preferred_theme !== undefined) {
      auth.user.preferred_theme = payload.preferred_theme === "contrast" ? "contrast" : "dark";
    }
    if (payload.last_route !== undefined) {
      auth.user.last_route = normalizeRouteForRole(payload.last_route, auth.user.role);
    }
    auth.user.updated_at = nowIso();
    await persistStorage(db, env);
    return jsonResponse({ user: serializeUser(auth.user) });
  }

  if (url.pathname === "/api/auth/password" && request.method === "POST") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    const payload = await request.json();
    const currentPassword = String(payload.current_password || "");
    const newPassword = String(payload.new_password || "").trim();
    const confirmPassword = String(payload.confirm_password || "").trim();

    if (!auth.user.must_change_password && !(await verifyPassword(currentPassword, auth.user.password_hash))) {
      return jsonResponse({ error: "Senha atual invalida" }, 400);
    }
    if (newPassword.length < 4) {
      return jsonResponse({ error: "A nova senha deve ter pelo menos 4 caracteres" }, 400);
    }
    if (newPassword !== confirmPassword) {
      return jsonResponse({ error: "A confirmacao da senha nao confere" }, 400);
    }

    auth.user.password_hash = await hashPassword(newPassword);
    auth.user.password_plain = newPassword;
    auth.user.must_change_password = false;
    auth.user.updated_at = nowIso();
    await persistStorage(db, env);
    return jsonResponse({ ok: true, user: serializeUser(auth.user) });
  }

  if (url.pathname === "/api/dashboard/overview" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    return jsonResponse(buildOverview(db, auth.user, url));
  }

  if (url.pathname === "/api/analysis" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    return jsonResponse(buildAnalysis(db, auth.user, url));
  }

  if (url.pathname === "/api/history" && request.method === "GET") {
    const auth = await requireAuth(request, db, env);
    if (auth.error) return auth.error;
    return jsonResponse(buildHistory(db, auth.user, url));
  }

  if (url.pathname === "/api/admin/users" && request.method === "GET") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    return jsonResponse({ users: db.users.map(serializeUser) });
  }

  if (url.pathname === "/api/admin/users" && request.method === "POST") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const payload = await request.json();
    const required = ["full_name", "login", "role"];
    const missing = required.filter((field) => !String(payload[field] || "").trim());
    if (missing.length) {
      return jsonResponse({ error: `Campos obrigatorios: ${missing.join(", ")}` }, 400);
    }
    if (!["operator", "manager"].includes(payload.role)) {
      return jsonResponse({ error: "Perfil invalido" }, 400);
    }
    if (db.users.some((entry) => entry.login === String(payload.login).trim())) {
      return jsonResponse({ error: "Login ja cadastrado" }, 409);
    }
    const user = {
      id: nextId(db, "users"),
      full_name: String(payload.full_name).trim(),
      login: String(payload.login).trim(),
      password_hash: await hashPassword(DEFAULT_PASSWORD),
      password_plain: DEFAULT_PASSWORD,
      role: payload.role,
      platform_0800_id: String(payload.platform_0800_id || "").trim(),
      nuvidio_id: String(payload.nuvidio_id || "").trim(),
      must_change_password: true,
      is_active: true,
      preferred_theme: "dark",
      last_route: "overview",
      created_at: nowIso(),
      updated_at: nowIso(),
    };
    db.users.push(user);
    await persistStorage(db, env);
    return jsonResponse({ user: serializeUser(user), default_password: DEFAULT_PASSWORD }, 201);
  }

  if (url.pathname.startsWith("/api/admin/users/") && request.method === "PUT") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const userId = Number(url.pathname.split("/").pop());
    const user = db.users.find((entry) => entry.id === userId);
    if (!user) return jsonResponse({ error: "Usuario nao encontrado" }, 404);
    const payload = await request.json();
    if (payload.role && !["operator", "manager"].includes(payload.role)) {
      return jsonResponse({ error: "Perfil invalido" }, 400);
    }
    if (payload.login && db.users.some((entry) => entry.login === String(payload.login).trim() && entry.id !== user.id)) {
      return jsonResponse({ error: "Login ja cadastrado" }, 409);
    }
    const wantsActive = payload.is_active === undefined ? user.is_active : Boolean(payload.is_active);
    if (!wantsActive && user.id === auth.user.id) {
      return jsonResponse({ error: "Voce nao pode desativar o proprio usuario" }, 400);
    }
    user.full_name = String(payload.full_name || user.full_name).trim();
    user.login = String(payload.login || user.login).trim();
    user.role = payload.role || user.role;
    user.platform_0800_id = String(payload.platform_0800_id ?? user.platform_0800_id).trim();
    user.nuvidio_id = String(payload.nuvidio_id ?? user.nuvidio_id).trim();
    user.is_active = wantsActive;
    user.updated_at = nowIso();
    if (String(payload.password || "").trim()) {
      user.password_hash = await hashPassword(String(payload.password).trim());
      user.password_plain = String(payload.password).trim();
      user.must_change_password = true;
    }
    await persistStorage(db, env);
    return jsonResponse({ user: serializeUser(user) });
  }

  if (url.pathname.startsWith("/api/admin/users/") && request.method === "DELETE") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const userId = Number(url.pathname.split("/").pop());
    if (userId === auth.user.id) {
      return jsonResponse({ error: "Voce nao pode apagar o proprio usuario" }, 400);
    }
    const userIndex = db.users.findIndex((entry) => entry.id === userId);
    if (userIndex === -1) return jsonResponse({ error: "Usuario nao encontrado" }, 404);
    db.users.splice(userIndex, 1);
    db.dailyMetrics = db.dailyMetrics.filter((entry) => entry.user_id !== userId);
    db.qualityScores = db.qualityScores.filter((entry) => entry.user_id !== userId);
    await persistStorage(db, env);
    return jsonResponse({ ok: true });
  }

  if (url.pathname === "/api/admin/quality" && request.method === "POST") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const payload = await request.json();
    const userId = Number(payload.user_id);
    if (!userId || !String(payload.reference_month || "").trim()) {
      return jsonResponse({ error: "user_id e reference_month sao obrigatorios" }, 400);
    }
    let score = db.qualityScores.find((item) => item.user_id === userId && item.reference_month === payload.reference_month);
    if (!score) {
      score = {
        id: nextId(db, "qualityScores"),
        user_id: userId,
        reference_month: payload.reference_month,
        score: toFloat(payload.score),
        notes: String(payload.notes || "").trim(),
        created_at: nowIso(),
        updated_at: nowIso(),
      };
      db.qualityScores.push(score);
    } else {
      score.score = toFloat(payload.score);
      score.notes = String(payload.notes || "").trim();
      score.updated_at = nowIso();
    }
    await persistStorage(db, env);
    return jsonResponse({ quality: score });
  }

  if (url.pathname === "/api/admin/quality/template" && request.method === "GET") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    if (!isNode) {
      const operators = db.users.filter((user) => user.is_active && user.role === "operator");
      const csv = [
        "Nome do Operador,Monitoria 1,Monitoria 2,Monitoria 3,Monitoria 4",
        ...operators.map((user) => `"${String(user.full_name || "").replaceAll('"', '""')}",,,,`),
      ].join("\n");
      return new Response(csv, {
        status: 200,
        headers: {
          "content-type": "text/csv; charset=utf-8",
          "content-disposition": 'attachment; filename="modelo-monitoria.csv"',
        },
      });
    }
    const operators = db.users.filter((user) => user.is_active && user.role === "operator");
    const buffer = await generateQualityTemplate(operators);
    return new Response(buffer, {
      status: 200,
      headers: {
        "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "content-disposition": 'attachment; filename="modelo-monitoria.xlsx"',
      },
    });
  }

  if (url.pathname === "/api/admin/quality/import" && request.method === "POST") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const form = await request.formData();
    const file = form.get("file");
    const referenceMonth = String(form.get("reference_month") || "").trim();
    if (!(file instanceof File)) {
      return jsonResponse({ error: "Arquivo de monitoria nao enviado" }, 400);
    }
    if (!isValidMonthRef(referenceMonth)) {
      return jsonResponse({ error: "Informe o mes de referencia no formato AAAA-MM" }, 400);
    }
    const lowerName = file.name.toLowerCase();
    if (!lowerName.endsWith(".xlsx") && !lowerName.endsWith(".csv")) {
      return jsonResponse({ error: "Use um arquivo XLSX ou CSV para a monitoria" }, 400);
    }

    if (!isNode && lowerName.endsWith(".xlsx")) {
      return cloudQualityUnsupported();
    }

    const rows = lowerName.endsWith(".csv")
      ? parseCsv(await file.text()).map((row) => ({
          name: row.nome_do_operador || row.nome_operador || row.operador || row.nome || "",
          monitoria_1: row.monitoria_1 ?? "",
          monitoria_2: row.monitoria_2 ?? "",
          monitoria_3: row.monitoria_3 ?? "",
          monitoria_4: row.monitoria_4 ?? "",
        }))
      : await parseQualityWorkbook(file);
    const operators = db.users.filter((user) => user.is_active && user.role === "operator");
    const errors = [];
    let processed = 0;

    for (const row of rows) {
      try {
        const user = matchUserByName(operators, row.name);
        if (!user) throw new Error(`Operador nao encontrado: ${row.name}`);
        const monitorias = [row.monitoria_1, row.monitoria_2, row.monitoria_3, row.monitoria_4];
        const validMonitorias = monitorias.filter((value) => value !== null && value !== undefined && value !== "");
        for (const value of validMonitorias) {
          const numeric = toFloat(value);
          if (numeric < 0 || numeric > 100) throw new Error(`Monitoria fora da faixa 0-100 para ${row.name}`);
        }
        const scoreValue = validMonitorias.length
          ? Number((validMonitorias.reduce((sum, value) => sum + toFloat(value), 0) / validMonitorias.length).toFixed(2))
          : 0;

        let score = db.qualityScores.find((item) => item.user_id === user.id && item.reference_month === referenceMonth);
        if (!score) {
          score = {
            id: nextId(db, "qualityScores"),
            user_id: user.id,
            reference_month: referenceMonth,
            score: scoreValue,
            notes: "",
            monitoria_1: row.monitoria_1,
            monitoria_2: row.monitoria_2,
            monitoria_3: row.monitoria_3,
            monitoria_4: row.monitoria_4,
            created_at: nowIso(),
            updated_at: nowIso(),
          };
          db.qualityScores.push(score);
        } else {
          score.score = scoreValue;
          score.monitoria_1 = row.monitoria_1;
          score.monitoria_2 = row.monitoria_2;
          score.monitoria_3 = row.monitoria_3;
          score.monitoria_4 = row.monitoria_4;
          score.updated_at = nowIso();
        }
        processed += 1;
      } catch (error) {
        errors.push(error.message);
      }
    }

    await persistStorage(db, env);
    return jsonResponse({ processed, rejected: errors.length, errors: errors.slice(0, 20) });
  }

  if (url.pathname.startsWith("/api/admin/daily-metrics/") && request.method === "PUT") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const metricId = Number(url.pathname.split("/").pop());
    const metric = db.dailyMetrics.find((entry) => entry.id === metricId);
    if (!metric) return jsonResponse({ error: "Registro nao encontrado" }, 404);
    const payload = await request.json();
    metric.production = toInt(payload.production ?? metric.production);
    metric.calls_0800_approved = toInt(payload.calls_0800_approved ?? metric.calls_0800_approved);
    metric.calls_0800_rejected = toInt(payload.calls_0800_rejected ?? metric.calls_0800_rejected);
    metric.calls_0800_pending = toInt(payload.calls_0800_pending ?? metric.calls_0800_pending);
    metric.calls_0800_no_action = toInt(payload.calls_0800_no_action ?? metric.calls_0800_no_action);
    metric.calls_nuvidio_approved = toInt(payload.calls_nuvidio_approved ?? metric.calls_nuvidio_approved);
    metric.calls_nuvidio_rejected = toInt(payload.calls_nuvidio_rejected ?? metric.calls_nuvidio_rejected);
    metric.calls_nuvidio_no_action = toInt(payload.calls_nuvidio_no_action ?? metric.calls_nuvidio_no_action);
    metric.updated_at = nowIso();
    await persistStorage(db, env);
    return jsonResponse({ metric: { ...metric, effectiveness: calculateEffectiveness(metric) } });
  }

  if (url.pathname === "/api/admin/import/r2" && request.method === "POST") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    if (!env.IMPORTS_BUCKET || typeof env.IMPORTS_BUCKET.list !== "function") {
      return jsonResponse({ error: "Bucket R2 nao configurado para este ambiente." }, 400);
    }

    const listing = await env.IMPORTS_BUCKET.list({ limit: 200 });
    const objects = (listing.objects || [])
      .filter((item) => /\.(csv|xlsx)$/i.test(item.key))
      .sort((a, b) => new Date(b.uploaded || 0).getTime() - new Date(a.uploaded || 0).getTime());

    if (!objects.length) {
      return jsonResponse({ error: "Nenhum arquivo CSV ou XLSX encontrado no R2." }, 404);
    }

    const normalizedCandidate = objects.find((item) => /\.csv$/i.test(item.key) && /base|produc|resultado|consolid/i.test(item.key));
    const latest0800 = objects.find((item) => /0800/i.test(item.key));
    const latestNuvidio = objects.find((item) => /nuvidio/i.test(item.key));
    const toProcess = normalizedCandidate ? [normalizedCandidate] : [latest0800, latestNuvidio].filter(Boolean);
    const summary = {
      processedFiles: 0,
      processedRows: 0,
      rejectedRows: 0,
      files: [],
      skipped: [],
    };

    for (const item of toProcess) {
      const object = await env.IMPORTS_BUCKET.get(item.key);
      if (!object) {
        summary.skipped.push(`${item.key}: arquivo nao encontrado no bucket.`);
        continue;
      }

      if (/\.xlsx$/i.test(item.key)) {
        summary.skipped.push(`${item.key}: leitura automatica de XLSX no R2 ainda nao esta habilitada. Use CSV para a base operacional.`);
        continue;
      }

      const text = await object.text();
      const rows = parseCsv(text);
      if (!rows.length) {
        summary.skipped.push(`${item.key}: arquivo vazio.`);
        continue;
      }

      const schema = detectImportSchema(rows[0], item.key);
      let result;
      if (schema === "normalized") result = importNormalizedRows(db, db.users.filter((entry) => entry.is_active), rows, item.key);
      else if (schema === "0800") result = import0800Rows(db, db.users.filter((entry) => entry.is_active), rows, item.key);
      else if (schema === "nuvidio") result = importNuvidioRows(db, db.users.filter((entry) => entry.is_active), rows, item.key);
      else {
        summary.skipped.push(`${item.key}: formato nao reconhecido para importacao automatica.`);
        continue;
      }

      registerImportLog(db, auth.user.id, item.key, result.processed, result.rejected);
      summary.processedFiles += 1;
      summary.processedRows += result.processed;
      summary.rejectedRows += result.rejected;
      summary.files.push({
        key: item.key,
        schema,
        processed: result.processed,
        rejected: result.rejected,
      });
      result.errors.slice(0, 10).forEach((entry) => {
        summary.skipped.push(`${item.key} linha ${entry.row}: ${entry.error}`);
      });
    }

    await persistStorage(db, env);
    return jsonResponse(summary);
  }

  if (url.pathname === "/api/admin/import" && request.method === "POST") {
    const auth = await requireManager(request, db, env);
    if (auth.error) return auth.error;
    const form = await request.formData();
    const file = form.get("file");
    if (!(file instanceof File)) {
      return jsonResponse({ error: "Arquivo nao enviado" }, 400);
    }
    if (!file.name.toLowerCase().endsWith(".csv")) {
      return jsonResponse({ error: "No formato atual, use CSV para o teste local." }, 400);
    }
    const text = await file.text();
    const rows = parseCsv(text);
    if (!rows.length) return jsonResponse({ error: "A planilha esta vazia." }, 400);
    const required = [
      "data",
      "producao",
      "0800_aprovado",
      "0800_cancelada",
      "0800_pendenciada",
      "0800_sem_acao",
      "nuvidio_aprovada",
      "nuvidio_reprovada",
      "nuvidio_sem_acao",
    ];
    const first = rows[0];
    const missing = required.filter((field) => !(field in first));
    if (missing.length) {
      return jsonResponse({ error: `Colunas obrigatorias ausentes: ${missing.join(", ")}` }, 400);
    }
    const { processed, rejected, errors } = importNormalizedRows(db, db.users.filter((entry) => entry.is_active), rows, file.name);
    registerImportLog(db, auth.user.id, file.name, processed, rejected);
    await persistStorage(db, env);
    return jsonResponse({ processed, rejected, errors: errors.slice(0, 20) });
  }

  return jsonResponse({ error: "Rota nao encontrada" }, 404);
}

async function serveStatic(request, env = {}) {
  const pathname = new URL(request.url).pathname;
  if (!isLocalNodeRuntime) {
    if (pathname === "/" || pathname === "/index.html") {
      return new Response(INDEX_HTML, {
        status: 200,
        headers: { "content-type": "text/html; charset=utf-8" },
      });
    }
    if (pathname === "/script.js") {
      return new Response(SCRIPT_JS, {
        status: 200,
        headers: { "content-type": "application/javascript; charset=utf-8" },
      });
    }
    if (pathname === "/styles.css") {
      return new Response(STYLES_CSS, {
        status: 200,
        headers: { "content-type": "text/css; charset=utf-8" },
      });
    }
    if (pathname === "/logos_KR-02.png") {
      return new Response(decodeDataUriBase64(FAVICON_DATA), {
        status: 200,
        headers: {
          "content-type": "image/png",
          "cache-control": "public, max-age=86400",
        },
      });
    }
    return new Response("Not found", { status: 404 });
  }
  const { fs, path } = await nodeModules();
  const directFile = STATIC_FILES[pathname];
  const fallbackFile = pathname.startsWith("/") ? pathname.slice(1) : pathname;
  let fileName = directFile;
  if (!fileName) {
    const candidate = path.resolve(fallbackFile);
    try {
      const stat = await fs.stat(candidate);
      if (stat.isFile()) fileName = fallbackFile;
    } catch {
      fileName = null;
    }
  }
  if (!fileName) {
    return new Response("Not found", { status: 404 });
  }
  const content = await fs.readFile(path.resolve(fileName));
  const type = contentType(fileName);
  return new Response(content, { status: 200, headers: { "content-type": type } });
}

function contentType(fileName) {
  if (fileName.endsWith(".html")) return "text/html; charset=utf-8";
  if (fileName.endsWith(".css")) return "text/css; charset=utf-8";
  if (fileName.endsWith(".js")) return "application/javascript; charset=utf-8";
  if (fileName.endsWith(".png")) return "image/png";
  return "application/octet-stream";
}

async function handleRequest(request, env = {}) {
  const url = new URL(request.url);
  if (url.pathname.startsWith("/api/")) {
    const db = await ensureStorage(env);
    return handleApi(request, url, db, env);
  }
  return serveStatic(request, env);
}

export default {
  fetch: async (request, env) => {
    try {
      return await handleRequest(request, env);
    } catch (error) {
      console.error("Worker error", error);
      const url = new URL(request.url);
      if (url.pathname.startsWith("/api/")) {
        return jsonResponse(
          {
            error: "Erro interno do worker",
            details: error?.stack || error?.message || String(error),
          },
          500,
        );
      }
      return new Response(
        `Erro interno do worker: ${error?.message || String(error)}`,
        {
          status: 500,
          headers: { "content-type": "text/plain; charset=utf-8" },
        },
      );
    }
  },
};

if (isLocalNodeRuntime) {
  const modules = await nodeModules();
  const __filename = modules.fileURLToPath(import.meta.url);
  if (process.argv[1] === __filename) {
    const server = modules.http.createServer(async (req, res) => {
      const bodyChunks = [];
      req.on("data", (chunk) => bodyChunks.push(chunk));
      req.on("end", async () => {
        const body = Buffer.concat(bodyChunks);
        const origin = `http://${req.headers.host}`;
        const request = new Request(new URL(req.url, origin), {
          method: req.method,
          headers: req.headers,
          body: req.method === "GET" || req.method === "HEAD" ? undefined : body,
        });
        try {
          const response = await handleRequest(request, {});
          res.statusCode = response.status;
          response.headers.forEach((value, key) => {
            res.setHeader(key, value);
          });
          const arrayBuffer = await response.arrayBuffer();
          res.end(Buffer.from(arrayBuffer));
        } catch (error) {
          res.statusCode = 500;
          res.setHeader("content-type", "application/json; charset=utf-8");
          res.end(JSON.stringify({ error: "Erro interno do servidor", details: error.message }));
        }
      });
    });
    const port = Number(process.env.PORT || 8787);
    server.listen(port, () => {
      console.log(`Pulse Ops disponÃ­vel em http://localhost:${port}`);
    });
  }
}

