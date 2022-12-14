using XLSX, DataFrames, LinearAlgebra

function trabalho()
    df = DataFrame(XLSX.readtable("FECHAMENTO_MAIS__NEGOCIADAS_5minutos.xlsx", "Plan1"))
    nr = nrow(df)
    cont = 1
    ninv = 0
    while cont <= 42-11
        try
            m = Matrix{Float64}(df[cont:cont+11, 2:13])
            if !isapprox(det(BigFloat.(m)), 0, atol = 1e-20)
                invm = inv(m)
                print("$invm\n\n")
            else 
                print("Determinante da matriz é zero ou muito perto de zero!\n\n")
                ninv += 1
            end
        catch err
            if isa(err, ArgumentError)
                print("Matriz não tem dados em todas celúlas!\n\n")
            elseif isa(err, MethodError)
                print("Matriz possui dados NAN!\n\n")
            elseif isa(err, SingularException)
                print("Não é possível inverter a matriz pois ela é singular!\n\n")
                ninv += 1
            else
                print("Erro: $err\n\n")
            end
        end
        cont += 1
    end
    print("\n\n$ninv\n\n")
end

@timev trabalho()