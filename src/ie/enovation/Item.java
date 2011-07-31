/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package ie.enovation;

import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author pvillega
 */
public class Item {
    private List<Metadata> metadata;

    public Item() {
        metadata = new ArrayList<Metadata>();
    }

    public void add(String dc, String value, String type) {
        if(type == null || "single".equals(type)){
            add(dc, value);
        }
        else if("multiple".equals(type)){
            addMultiple(dc, value);
        }
    }

    public void add(String dc, String value) {
        if(dc.split("\\.").length == 2){
            dc = dc+ ".none";
        }
        String[] key = dc.split("\\.");
        Metadata m = new Metadata();
        m.schema = key[0];
        m.element = key[1];
        m.qualifier = key[2];
        m.value = value;
        metadata.add(m);
    }

    public void addMultiple(String dc,String data){
        String[] vals = data.split(";");
        for(String a : vals){
            add(dc,a);
        }
    }

    //TODO: could be improved to store the language

    public String toXML(){
        StringBuilder bf = new StringBuilder();
        bf.append("<dublin_core>");

        //we iterate over the metadata        
        for(Metadata m: metadata){
            bf.append("<dcvalue element=\"").append(m.element).append("\" ");
            bf.append("qualifier=\"").append(m.qualifier).append("\">");
            bf.append(m.value);
            bf.append("</dcvalue>");
        }

        bf.append("</dublin_core>");
        return bf.toString();
    }
}
