library("data.table")
library(dplyr)
library(ggplot2)
library(patchwork)
library(stats)

#=============================================================

rm(list=ls())
setwd("C:/Users/Admin/Documents/DS-Docs/Proj1")

dfi <- fread("PI2017_small_50.txt", encoding="UTF-8", dec = ";")

df <- select(dfi, `Escolaridade após 2005`, `Vl Remun Média Nom`, `Raça Cor`)


#=============================================================

#CHARACTER TO NUMERIC (TRANSF COMMA TO POINT)
df$`Vl Remun Média Nom` <- as.numeric(sub(",", ".", df$`Vl Remun Média Nom`, fixed = TRUE))

#CREATED DF0 -- SALARY > 0

df0 <- filter(df, `Vl Remun Média Nom` > 0)

#CREATED DF69 -- Self-Declared Color is not null

df69 <- filter(df, `Raça Cor` < 9)

#CREATED DF70 -- SALARY > 0 & Self-Declared Color is not null

df70 <- filter(df69, `Vl Remun Média Nom` > 0)

#FACTOR: Education Level & Self-Declared Color

df$`Escolaridade após 2005` <- as.factor(df$`Escolaridade após 2005`)
df0$`Escolaridade após 2005` <- as.factor(df0$`Escolaridade após 2005`)
df69$`Escolaridade após 2005` <- as.factor(df69$`Escolaridade após 2005`)
df70$`Escolaridade após 2005` <- as.factor(df70$`Escolaridade após 2005`)
df$`Raça Cor` <- as.factor(df$`Raça Cor`)
df0$`Raça Cor` <- as.factor(df0$`Raça Cor`)
df69$`Raça Cor` <- as.factor(df69$`Raça Cor`)
df70$`Raça Cor` <- as.factor(df70$`Raça Cor`)

#=============================================================

# PLOTTING

#=============================================================

# 100% Stacked Bar Chart Self-Declared Color vs Education Level
df69 %>%
  ggplot(df69, mapping = aes(x = `Escolaridade após 2005`, fill = `Raça Cor`)) +
  geom_bar(position = "fill") +
  theme_classic()+
  scale_fill_discrete(name = "Self-Declared Color", labels = c("Indigenous", "White", "Black", "Yellow", "Brown"))+
  labs(x = "Education Level", y = "Density")

#=============================================================

# Boxplots (without outliers) - Self-Declared Color vs Salary

df70 %>%
  ggplot(df70, mapping = aes(x = `Raça Cor`, y = `Vl Remun Média Nom`, fill = `Raça Cor`, group = `Raça Cor`)) +
  geom_boxplot(outlier.shape = NA) +
  coord_cartesian(ylim=c(0,4000)) +
  scale_fill_discrete(name = "Self-Declared Color", labels = c("Indigenous", "White", "Black", "Yellow", "Brown"))+
  labs(x = "Self-Declared Color", y = "Average Monthly Salary (BRL)")

#=============================================================

# Dispersion with Jitter Points - Education Level vs Salary
df0 %>%
  ggplot(df0, mapping = aes(x = factor(`Escolaridade após 2005`), y = `Vl Remun Média Nom`, color = `Escolaridade após 2005`)) +
  scale_color_discrete(name = "Education Level", labels = c("Illiterate", "Up to 4th grade Incompl.", "Primary School Degree", "5th to 8th grade Incompl.", "Lower Secondary School Degree", "Up to 11th grade", "High School Degree", "Undergrad. Incompl.", "Undergrad. Degree", "Master's Degree", "Doctorate Degree")) +
  geom_jitter(size=1) +
  coord_cartesian(ylim=c(0,80000)) +
  scale_y_continuous(labels=function(n){format(n, scientific = FALSE)}) +
  theme_classic() +
  labs(x = "Education Level", y = "Average Monthly Salary (BRL)")

#=============================================================


# Salary means grouped by Self-Declared Color

mean(df70$`Vl Remun Média Nom`[df70$`Raça Cor` == 1])
mean(df70$`Vl Remun Média Nom`[df70$`Raça Cor` == 2])
mean(df70$`Vl Remun Média Nom`[df70$`Raça Cor` == 4])
mean(df70$`Vl Remun Média Nom`[df70$`Raça Cor` == 6])
mean(df70$`Vl Remun Média Nom`[df70$`Raça Cor` == 8])

#=============================================================

# Salary means grouped by Education Level

mean(df0$`Vl Remun Média Nom`[df0$`Escolaridade após 2005` == 1])
mean(df0$`Vl Remun Média Nom`[df0$`Escolaridade após 2005` == 2])
mean(df0$`Vl Remun Média Nom`[df0$`Escolaridade após 2005` == 3])
mean(df0$`Vl Remun Média Nom`[df0$`Escolaridade após 2005` == 4])
mean(df0$`Vl Remun Média Nom`[df0$`Escolaridade após 2005` == 5])
mean(df0$`Vl Remun Média Nom`[df0$`Escolaridade após 2005` == 6])
mean(df0$`Vl Remun Média Nom`[df0$`Escolaridade após 2005` == 7])
mean(df0$`Vl Remun Média Nom`[df0$`Escolaridade após 2005` == 8])
mean(df0$`Vl Remun Média Nom`[df0$`Escolaridade após 2005` == 9])
mean(df0$`Vl Remun Média Nom`[df0$`Escolaridade após 2005` == 10])
mean(df0$`Vl Remun Média Nom`[df0$`Escolaridade após 2005` == 11])

#============================================================

# Normal Distributions - Salary grouped by Self-Declared Color

ggplot(df70, aes(x=`Vl Remun Média Nom`, color=`Raça Cor`)) +
  geom_density(size=1)+
  labs(y = "Density", x = 'Average Monthly Salary (BRL)') +
  scale_color_discrete(name = "Self-Declared Color", labels = c("Indigenous", "White", "Black", "Yellow", "Brown"))+
  coord_cartesian(xlim=c(0,5000))

#============================================================

# CREATED DF WHERE Self-Declared Color == White | Black
dft <- filter(df0, `Raça Cor` == '2' | `Raça Cor` == '4')

#Welch
t.test(`Vl Remun Média Nom`~`Raça Cor`, data=dft)

#Student - not used
#t.test(`Vl Remun Média Nom`~`Raça Cor`, data=dft, var.equal=TRUE)

#====================== LORENZ CURVE =======================================
#--------------- WRITTEN BY BRUNO NEGREIROS --------------------------------

renda = df0$ `Vl Remun Média Nom`
quantile(renda, probs = seq(0,1,0.1) )
classe = cut(renda, 
             breaks = quantile(renda, probs = seq(0,1,0.1)),
             include.lowest = TRUE,
             labels=FALSE)

#DF of 2 main variables created
tabela <- data.frame(renda, classe)

#Absolute frequency of population and mass salary grouped by quartiles
tabela_gini <- tabela %>% group_by(classe) %>%
  summarise(pop_tot = n(),
            renda_tot = sum(renda)) %>% ungroup()
#Relative frequencies calculated
tabela_gini <- tabela_gini %>% mutate (freq_pop = pop_tot/sum(pop_tot),
                                       freq_renda = renda_tot/sum(renda_tot))
#Null line created
linha_add = rep(0, ncol(tabela_gini))

#DFs named
names(linha_add) <- names(tabela_gini)

#DFs merged with null line
tabela_gini <- bind_rows(linha_add, tabela_gini)

#=============================================================
# Cumulative Frequency Created = Unitary interval
#=============================================================

tabela_gini <- tabela_gini %>% mutate(freq_pop_ac = cumsum(freq_pop),
                                      freq_renda_ac = cumsum(freq_renda))

#=============================================================
# Lorenz Curve Created
#=============================================================
plot_lorenz <- ggplot(data = tabela_gini, aes(x=freq_pop_ac, y=freq_renda_ac)) +
  geom_line(color="red", size = 1.25) +
  geom_point(size=2, color="red") +
  xlim(0, 1) +
  ylim(0, 1) +
  labs(x="Population Accumulated", y="Average Monthly Salary Accumulated")
plot_lorenz

#=============================================================
# Equality line created
#=============================================================
plot_lorenz +
  geom_line(aes(x=freq_renda_ac), color="blue", size= 1 )+
  coord_fixed()

#=============================================================
# Gini Coefficient Calculated
#=============================================================
tabela_gini <- tabela_gini %>% mutate(altura = freq_pop_ac - lag(freq_pop_ac),
                                      media = (freq_renda_ac + lag(freq_renda_ac))/2,
                                      area = media*altura
)
beta = sum(tabela_gini$area, na.rm = TRUE)
beta

gini = 2*(.5 - beta)
gini