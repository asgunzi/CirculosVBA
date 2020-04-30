# CirculosVBA

Desenho de padrões com círculos - em VBA

Quando eu era adolescente, vi uma espécie de régua, que permitia fazer padrões bonitos a partir de elipses.

Era um negócio desse tipo:
![](https://ferramentasexcelvba.files.wordpress.com/2020/04/spirograph.jpg?w=1024)


Dá para fazer algo mais ou menos parecido, com apenas um comando VBA: o de plotar círculos. O formato é o seguinte:

`ActiveSheet.Shapes.AddShape(msoShapeOval, x, y, comprim, largura)`

Para plotar um círculo, devemos saber a posição dele (x,y), o comprimento e largura. Se é um círculo e não uma elipse, comprimento = largura = raio.

O resto dos comandos serve apenas para apagar desenhos antigos, colorir e posicionar os novos círculos.

Algumas convenções:

    Número de círculos é o número de círculos a plotar.

    O raio menor é o raio desses círculos.

    O raio maior é o raio em torno da qual os círculos são plotados

    E o ângulo, em graus, é o ângulo entre um círculo e outro.

A ilustração a seguir dá uma ideia desses parâmetros.
![](https://ferramentasexcelvba.files.wordpress.com/2020/04/circulos01.png)


Obs. Como o ângulo é divisor de 360 graus, tem hora que os círculos ficam um sobre o outro.

Com apenas esses conceitos, é possível gerar algumas variações interessantes.

![](https://ferramentasexcelvba.files.wordpress.com/2020/04/circulos02.png)

![](https://ferramentasexcelvba.files.wordpress.com/2020/04/circulos03.png)

Planilha para download no Github. Para rodar, é necessário ativar macros.

![](https://ferramentasexcelvba.files.wordpress.com/2020/04/circulos04.png)

Outras sugestões são bem-vindas.

[https://ideiasesquecidas.com/](https://ideiasesquecidas.com/)

[https://ferramentasexcelvba.wordpress.com/2020/04/30/desenho-de-padroes-com-circulos/](https://ferramentasexcelvba.wordpress.com/2020/04/30/desenho-de-padroes-com-circulos/)


