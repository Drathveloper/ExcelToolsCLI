package org.drathveloper.excelToolsCLI.CLIUtils;

import org.apache.commons.lang3.ArrayUtils;

import java.util.*;

public class ArgsParser {

    private Map<String, List<String>> arguments;

    private ArgsParser(String[] arguments){
        this.arguments = this.processArguments(arguments);
    }

    public ArgsParser(String[] arguments, String[] requiredOptions) throws IllegalArgumentException {
        this(arguments);
        if(!this.areAllOptionsInList(requiredOptions)){
            throw new IllegalArgumentException("Not enough parameters");
        } else {
            this.arguments = this.processArguments(arguments);
        }
    }

    public ArgsParser(String[] arguments, String[] requiredOptions, String[] optionalOptions) throws IllegalArgumentException {
        this(arguments);
        if (!this.areAllOptionsInList(requiredOptions)) {
            throw new IllegalArgumentException("Not enough parameters");
        } else if(!this.areAllLegalOptions(requiredOptions, optionalOptions)){
            throw new IllegalArgumentException("Illegal arguments");
        }else {
            this.arguments = this.processArguments(arguments);
        }
    }

    private Map<String, List<String>> processArguments(String[] arguments){
        Map<String, List<String>> params = new HashMap<>();
        String auxParam = null;
        List<String> auxList = new ArrayList<>();
        for(String param : arguments){
            if(param.charAt(0) == '-'){
                if(auxParam!=null){
                    params.put(auxParam, new ArrayList<>(auxList));
                    auxList = new ArrayList<>();
                }
                auxParam = param;
            } else {
                auxList.add(param);
            }
        }
        if(auxParam!=null){
            params.put(auxParam, new ArrayList<>(auxList));
        }
        return params;
    }

    public List<String> getParametersFromOption(String option){
        if(arguments.get(option)!=null){
            return arguments.get(option);
        }
        return new ArrayList<>();
    }

    public String getParametersFromOptionAsString(String option){
        List<String> parameters = this.getParametersFromOption(option);
        return this.getArrayParameterAsString(parameters);
    }

    public boolean isOptionInArgs(String option){
        return arguments.get(option)!=null;
    }

    private String getArrayParameterAsString(List<String> list){
        StringBuilder output = new StringBuilder();
        for(int i=0; i<list.size(); i++){
            output.append(list.get(i));
            if(i < list.size() - 1){
                output.append(", ");
            }
        }
        return output.toString();
    }

    private boolean areAllLegalOptions(String[] requiredOptions, String[] optionalOptions){
        String[] allOptions = ArrayUtils.addAll(requiredOptions, optionalOptions);
        List<String> list = Arrays.asList(allOptions);
        for(String option : arguments.keySet()){
            if(!list.contains(option)){
                return false;
            }
        }
        return true;
    }

    private boolean areAllOptionsInList(String[] expectedOptions){
        int foundOptions = 0;
        for(String option : expectedOptions){
            if(this.isOptionInList(option)){
                foundOptions++;
            }
        }
        return foundOptions == expectedOptions.length;
    }

    private boolean isOptionInList(String option) {
        return arguments.keySet().contains(option);
    }

}
